#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Sep 15 09:06:02 2025

@author: Tjebbe Janson (RIVM)

Dit script maakt van gebiedsdossiers (pdf) vector bestanden waarbij de semantische betekenis van de tekst wordt vastgelegd
Vervolgens worden bij een vraagstelling relevante paragraven gezocht met cosinus gelijknis om uiteindelijk voor te leggen aan een LLM

Het taalmodel (LLM) kan met de extra context de vraag beter beantwoorden, deze techniek heet ook wel retrieval augmented generation (RAG)

Deel 1: parsen van documenten
Deel 2: inlezen aangepaste document bestanden en bevraag LLM

-- 

Script to parse documents and use AI to question the parsed files

"""

# =============================================================================
# Functions
# =============================================================================

import pdfplumber
import re
import os
import openai
from openai import OpenAI
import json
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np
import time
import pandas as pd
import camelot
import openpyxl
import warnings
import tiktoken
import copy
from collections import Counter
import base64
from pdf2image import convert_from_path
from pypdf import PdfReader


# Filepaths 
# Retrieve_table() slaat hier afbeeldingen op van pagina's met tabellen. Deze tabellen 
quickstorage_pdf_images = ""

# Hier worden uiteindelijk de vectoren per gebiedsdossier opgeslagen
quickStorage = ""
vectorized_result_directory = ""

# Hier staan de PDF bestanden van gebiedsdossiers
data_directory = ""

#Result filepaths hier schrijft 
correct = pd.read_excel("/results/correct.xlsx", sheet_name='Sheet1')
result_directory = "/results"

#settings
embedding_model = ["text-embedding-ada-002", "text-embedding-3-large"]
model_n = 1 #model keuze, index van lijst
n = 7 #iterations of opkomende stoffen


# LLM credentials en clients
key = ""
endpoint = ""

client = openai.AzureOpenAI(
    api_key=key,
    api_version="2024-02-01",
    azure_endpoint=endpoint
)
client_llm = openai.AzureOpenAI(
    api_key=key,
    api_version="2023-05-15",
    azure_endpoint=endpoint)

url = f"{endpoint}/openai/v1/"
client_4o = OpenAI(
    api_key=key,
    base_url=url,
)
# Referentie(s) Bijlage(s/n) definities begrippenlijst

# zoek naar TOC, inhoudsopgave uit bestand zelf
def extract_outline_as_list(filepath, blank_pages=0):
    reader = PdfReader(filepath)
    outline = reader.outline
    result = []

    def walk_outline(outline, prefix=""):
        for item in outline:
            if isinstance(item, list):
                walk_outline(item, prefix)
            else:
                title = str(item.title)
                # Bepaal het 'paragraph_number' uit de titel (optioneel, afhankelijk van jouw structuur)
                # Bijvoorbeeld als de titel begint met "1.2" of "Hoofdstuk 1"
                paragraph_number = None
                parts = title.strip().split(" ", 1)
                if parts[0][0].isdigit():
                    paragraph_number = parts[0]
                else:
                    paragraph_number = ""  # Of iets anders

                # Pagina ophalen (en +1 want meestal 0-based)
                if hasattr(item, "page") and item.page is not None:
                    page_number = reader.get_destination_page_number(item) + 1 + blank_pages
                else:
                    page_number = None

                result.append({
                    'paragraph': paragraph_number.replace('%',''),
                    'title': title.replace('%',''),
                    'start_page': page_number
                })

    walk_outline(outline)
    for i, item in enumerate(result):
        if i < len(result) - 1:
            item['end_page'] = result[i+1]['start_page']
        else:
            # Laatste hoofdstuk: tot het eind van het document
            item['end_page'] = len(reader.pages)
    return result

# Extract inhoudsopgave uit PDF met regular expression
def TOC(pdf, blank_pages=0):
    print("reconstructing TOC..")
    TOC = []
    # with pdfplumber.open(pdf) as pdf:
    toc_text = ""
    for page in pdf.pages[0:13]:
        toc_text += page.extract_text()
    parts = re.split(r'Inhoudsopgave|Inhoud', toc_text, flags=re.IGNORECASE)
    try:
        #pattern = r'^\s*\d+(?:\.\d+)*(?:\.+)?\s+.+?\s+\d+(?=\s|$)'
        pattern = r'^\s*\d+(?:\.\d+)*(?:\.+)?\s+.+?[\s\-–—]+(\d+)\s*$'
        #pattern2 = r'^\s*(\d+(?:\.\d+)*\.?)\s+(.+?)\s+(\d+)\s*$'
        pattern2 = r'^\s*(\d+(?:\.\d+)*\.?)\s+(.+?)[\s\-–—]+(\d+)\s*$'
        eind_pattern = r'\b(referenties|bijlage?: s/n?|definities|begrippenlijst)\b'
        part = parts[1]
        if len(part) < 500:
            part = parts[2]
        for i, line in enumerate(part.split("\n")):
            if i == 54:
                break
            if re.match(pattern, line):
                
                match = re.match(pattern2, line)
    
                if match:
                    paragraph_number = match.group(1)
                    title = match.group(2).replace(".", "")
                    page_number = int(match.group(3))
                    TOC.append({
                        'paragraph': paragraph_number,
                        'title': title,
                        'start_page': page_number + blank_pages
    
                    })
                    if len(TOC) > 1:
                        TOC[-2]['end_page'] = page_number + blank_pages
            match_eind = re.search(eind_pattern, line, re.IGNORECASE)
            if match_eind:
                match_pagenumber = re.findall(r'\d+$', line)
                if match_pagenumber:
                    if (int(match_pagenumber[0]) > int(TOC[-1]['start_page'])) & ((int(match_pagenumber[0]) - int(TOC[-1]['start_page'])) < 7):
                        TOC[-1]['end_page'] = int(match_pagenumber[0])
    except:
        pass
    # omgaan met bijlages en einde van document, bijlages niet in toc
    # Als laatste item geen eind pagina heeft, deze zelf geven om te lange chunk te vermijden
    # Indien geen TOC in document, pass
    
    try:
        if 'end_page' not in TOC[-1].keys():
            if len(pdf.pages) < int(TOC[-1]['start_page'] + 4):
                TOC[-1]['end_page'] = len(pdf.pages)
            else:
                TOC[-1]['end_page'] = int(TOC[-1]['start_page'] + 4)
    except:  
        pass
    if (TOC == []) | (len(TOC) < 5):
        try:
        #first pdf meta data:
            TOC = extract_outline_as_list(pdf.filepath)
        except:
            pass
    
    return TOC

# Vind de pagina waar het document start, soms lege pagina's (bijv. Titelpagina)
def find_start_page(pdf):

    for i, page_obj in enumerate(pdf.pages[0:13]):

        page = page_obj.extract_text()

        last_line = page.split("\n")[-1]
        match1 = re.search(r'(\d+)\s+van\s+\d+', last_line)  # 3 van 26
        match2 = re.search(r'(\d+)\s*$', last_line)  # getal rechterkant

        start_page_count = 0
        if match1:
            if int(match1.group(1)) > 13:
                continue
            if int(match1.group(1)) != i+1:
                start_page_count = i
                break
            else:
                start_page_count = 0

        elif match2:
            if int(match2.group(1)) > 13:
                continue
            if int(match2.group(1)) != i+1:
                start_page_count = i
                break
            else:
                start_page_count = 0
    return start_page_count

#Functie om obv pagina nummers uit TOC() een chunk text te halen, hier tabellen verwerken uit retrieve_table()
def get_chunk(pdf, toc, paragraph_index, tables, page_based=True):
    print(f'\t {paragraph_index} : {toc[paragraph_index]["title"]}')
    start_page = toc[paragraph_index]['start_page']
    end_page = toc[paragraph_index]['end_page']
    page_obj = pdf.pages[start_page-1:end_page]
    chunk = ""
    if page_based:
        for page in page_obj:
            # print(page.extract_text())
            chunk += page.extract_text()
    else:
        text = ""
        for page in page_obj:
            text += page.extract_text()
        begin_index = 'empty'
        eind_index = 'empty'
        table_regex = r'tabel\s+\d+([.-]\d+)?'
        table_line_index = []
        for index, line in enumerate(text.split("\n")):
            if re.match(table_regex, line, re.IGNORECASE):
                # if 'tables' not in locals().keys():
                #     print("creating tables...")
                #     tables = retrieve_table(filepath, paragraph_index)
                #     print("done.")
                table_line_index.append(index)
            length_pdf_head_line = len(
                line.strip().lower().replace(" ", "").replace(".", ""))
            if length_pdf_head_line < 10:
                length_pdf_head_line = 10
            if line.strip().lower().replace(" ", "").replace(".", "") == (toc[paragraph_index]['paragraph'] + toc[paragraph_index]['title']).strip().lower().replace(" ", "").replace(".", "")[0:length_pdf_head_line]:
                # Direct match:

                begin_index = index
            if begin_index == 'empty':
                continue
            if paragraph_index == len(toc)-1:
                # Last paragraph requested:
                eind_index = len(text.split("\n"))
                break
            elif line.strip().lower().replace(" ", "").replace(".", "") == (toc[paragraph_index+1]['paragraph'] + toc[paragraph_index+1]['title']).strip().lower().replace(" ", "").replace(".", "")[0:length_pdf_head_line]:
                # Direct match
                eind_index = index

        if eind_index == 'empty':
            eind_index = len(text.split("\n"))
        lines = text.split("\n")

        if (tables != {}) & (table_line_index != []):
            matching_keys = []
            for page in range(start_page, end_page+1):
                matching_keys.extend(
                    [key for key in tables.keys() if key[0] == page])
            # if len(table_line_index) < len(matching_keys):
            #     tables_combine =
            if (len(matching_keys) != 0):
                for num, line_index in enumerate(table_line_index):
                    try:
                        lines.insert(line_index+1, "\n" +
                                     tables[matching_keys[num]])
                    except Exception as e:
                        print(f"\t\t{e}")

        lines = lines[begin_index:eind_index]
        chunk = ' '.join(lines)
        encoding = tiktoken.encoding_for_model("gpt-3.5-turbo")
        if len(encoding.encode(chunk)) > 8190:
            print(
                f"\t\t[get_chunk]: Amount of tokens of chunk: {len(encoding.encode(chunk))}, is too high for embedding model\n\t\tremoving markdown tables...")
            lines = text.split("\n")
            lines = lines[begin_index:eind_index]
            chunk = ' '.join(lines)
    return chunk

# Functie om JSON te maken met per TOC item: id, dossier, paragraph, paragraph nr, resultaat get_chunk(), per chunk embeddings maken
def vectorize_json(toc, dossier, id_number, pdf, tables, model, page_based=True):
    json = []
    print(dossier)
    for i, paragraph in enumerate(toc):
        json.append({
            'id': id_number,
            'dossier': dossier.replace('.pdf', ''),
            'paragraph': paragraph['title'],
            'paragraph_number': paragraph['paragraph'],
            'content': get_chunk(pdf, toc, i, tables, page_based=page_based)
        })
    content = [item['content'] for item in json]
    paragraphs = [item['paragraph'] for item in json]
    for i, item in enumerate(json):
        # text-embedding-3-large
        response_paragraphs = client.embeddings.create(
            input=paragraphs[i], model=model)
        response_content = client.embeddings.create(
            input=content[i], model=model)
        paragraph_embeddings = response_paragraphs.data[0].embedding
        content_embeddings = response_content.data[0].embedding
        item['paragraph_vector'] = paragraph_embeddings
        item['content_vector'] = content_embeddings

    return json

# Query voor cosinus gelijknis embedden
def embed_query(string, model):
    query = client.embeddings.create(input=string, model=model)
    return query.data[0].embedding

# cosine gelijkenis runnen voor contents en query
def cosine_similarity_func(json_var, query, paragraven_n=5, print_score=False):
    # Removing smaller contents, headers etc
    json_temp = copy.deepcopy(json_var)
    for index, item in enumerate(json_var):

        if len(json_var[item]['content']) < 75:
            del json_temp[item]

    paragraphs_vectors = np.array(
        [json_temp[p]["content_vector"] for p in json_temp])
    paragraphs_text = [json_temp[p]['content'] for p in json_temp]
    similarity = cosine_similarity(paragraphs_vectors, [np.array(query)])
    similarity_scores = similarity.flatten()  # van (51, 1) naar (51,)
    top3_indices = np.argsort(similarity_scores)[-paragraven_n:][::-1]
    all_indices = np.argsort(similarity_scores)[::-1]
    scores_top = np.sort(similarity_scores)[::-1]
    if print_score:
        for index, score in zip(all_indices, scores_top):
            index = [x for x in json_var.keys()][index]
            print(str(json_var[index]['paragraph_number']) + "\t" + str(
                json_var[index]['paragraph']) + " :\n\t" + str(float(score)) + "\n")

    text = []
    for i, idx in enumerate(top3_indices, 1):
        # print(f"\n{i}. Score: {similarity_scores[idx]:.3f}")
        # print(paragraphs_text[idx])
        text.append(paragraphs_text[idx])
    return text

# functie alle tabellen uit pdf te halen en index mee te geven
def retrieve_table(client_4o, filepath):
    # print("pIndex: " + str(paragraph_index))
    print("reconstructing tables...")
    quickStorage = quickstorage_pdf_images
    table_regex = r'tabel\s+\d+([.-]\d+)?'
    text = ''
    table_pages_list = []
    table_pages_string = ''
    tables = {}
    # Regex om tabellen te vinden
    with pdfplumber.open(filepath) as pdf:
        total_length = len(pdf.pages)
        for index, page in enumerate(pdf.pages):

            text = page.extract_text()
            for line in text.split('\n'):
                if re.match(table_regex, line, re.IGNORECASE):
                    # print(line)
                    if index + 1 == total_length:
                        # table_pages_string += f', {index+1}'
                        table_pages_list.append([index+1])
                        continue
                    # print(index)
                    table_pages_list.append([index+1, index+2])
    print(f"{len(table_pages_list)} pages found with tables...")
    # table_pages_string += f', {index+1}, {index+2}' #extra pagina voor lange tabellen
    # table_pages_string = table_pages_string[2:]
    if table_pages_list == []:
        return {}
    uniques = []
    seen = set()
    for sublist in table_pages_list:
        t = tuple(sublist)
        if t not in seen:
            uniques.append(sublist)
            seen.add(t)
    # Loop door unieke hits om van pagina image op te slaan
    for i, setx in enumerate(uniques):
       
        # print(setx)
        if len(setx) == 2:
            filepaths = [f'{filepath.split("/")[-1].replace(".pdf", "")}_pagina_{setx[0]}.png',
                         f'{filepath.split("/")[-1].replace(".pdf", "")}_pagina_{setx[1]}.png']
        elif len(setx) == 1:
            filepaths = [
                f'{filepath.split("/")[-1].replace(".pdf", "")}_pagina_{setx[0]}.png']
        else:
            break
        indices = [x for x, y in enumerate(setx)]
        for x, file in enumerate(filepaths):
            if len(setx) == 0:
                break
            # Creating files
            if os.path.exists(quickStorage + file):
                indices.remove(x)
                continue
            else:
                if len(indices) == 2:
                    pages = convert_from_path(
                        filepath, 250, first_page=setx[0], last_page=setx[1], grayscale=True)
                    for page_n, page in zip(filepaths, pages):
                        page.save(quickStorage + page_n, 'PNG')
                    break
                elif len(indices) == 1:
                    pages = convert_from_path(
                        filepath, 250, first_page=setx[0], last_page=setx[0], grayscale=True)
                    for page_n, page in zip(filepaths, pages):
                        page.save(quickStorage + page_n, 'PNG')

        # Opening base64 image en bevraag LLM
        with open(quickStorage + filepaths[0], "rb") as f:
            image1 = base64.b64encode(f.read()).decode()
        if len(filepaths) == 2:
            with open(quickStorage + filepaths[1], "rb") as f:
                image2 = base64.b64encode(f.read()).decode()
            # Getting response
            try:
                response = client_4o.responses.create(
                    model="gpt-4o-research",
                    input=[
                        {
                            "role": "user",
                            "content": [
                                {"type": "input_text", "text": "Je bent een assistent die erg goed is in het vinden en interpreteren van tabellen in documenten. Kan je de tabel(len) op de pagina vinden en reconstrueren in markdown format? Volg deze regels nauwkeurig: 1) Als je een of meerdere tabellen vind geef je deze als markdown. 2) als je geen tabel vind, geef je een lege string. 3) behoudt de inhoud goed, zowel gevulde als lege celllen moeten zichtbaar zijn in de markdown tabel. 4) kijk goed naar bijschriften van tabellen, vind je meerdere tabellen die bij elkaar horen (1 bijschrift) voeg deze dan samen."},
                                {
                                    "type": "input_image",
                                    "image_url": f"data:image/jpeg;base64,{image1}"
                                },
                                {
                                    "type": "input_image",
                                    "image_url": f"data:image/jpeg;base64,{image2}"
                                }
                            ]
                        }
                    ]
                )
            except Exception as e:
                print(f"[retrieve_table] error while getting response: {e}")
        else:
            try:
                response = client_4o.responses.create(
                    model="gpt-4o-research",
                    input=[
                        {
                            "role": "user",
                            "content": [
                                {"type": "input_text", "text": "Je bent een assistent die erg goed is in het vinden en interpreteren van tabellen in documenten. Kan je de tabel(len) op de pagina vinden en reconstrueren in markdown format? Volg deze regels nauwkeurig: 1) Als je een of meerdere tabellen vind geef je deze als markdown. 2) als je geen tabel vind, geef je een lege string. 3) behoudt de inhoud goed, zowel gevulde als lege celllen moeten zichtbaar zijn in de markdown tabel. 4) kijk goed naar bijschriften van tabellen, vind je meerdere tabellen die bij elkaar horen (1 bijschrift) voeg deze dan samen."},
                                {
                                    "type": "input_image",
                                    "image_url": f"data:image/jpeg;base64,{image1}"
                                }
                            ]
                        }
                    ]
                )
            except Exception as e:
                print(f"[retrieve_table] error while getting response: {e}")
        tokens += response.usage.total_tokens
    # Returns markdown table constructed by LLM 
    return tables, tokens

# Vraag, cosine text etc opsturen naar LLM
def ask_llm(client, text, vraag, randvoorwaarde, toc, print_prompt=False, toelichting=False):

    prompt = (
        f"Hier is de inhoudsopgave en 5 relevante tekstfragmenten uit mijn document:\n\n"
        f"Inhoudsopgave: {toc}\n"
        f"5 relevante tekstfragmenteb\n{text}\n\n"
        f"Je bent een assistent die goed is in het zoeken van informatie uit dit document, het document beschrijft een drinkwaterwinning. Gebruik de inhoudsopgave om het juiste fragment te vinden."
        f"Volg goed de regels: 1) beantwoord de vraag met behulp van informatie uit het document. 2) volg strikt de regels uit de randvoorwaarde. 3) als je geen antwoord vind geef dan een lege string als resultaat. 4) voeg geen eigen toelichting of interpretatie toe. 5) tabellen staan weergeven als markdown, als het antwoord in deze markdown tabellen staat geef je deze informatie prioriteit. 6) haal de informatie uit de tabel en geef niet de hele tabel als resultaat."
        f"\nMijn vraag is:\n{vraag}\n\n"
        f"Geef een antwoord op basis van deze fragmenten en de volgende randvoorwaardes: {randvoorwaarde}.\n"
        f"Regels: Beantwoord de vraag en geef een resultaat wat strikt voldoet aan de randvoorwaardes, geef voorkeur aan gegevens die volgens de vraag en randvoorwaarde kloppen."
    )
    
    if toelichting:
        prompt += "\nVergeet de instructie om geen toelichting te geven, schrijf in een paar zinnen na het resultaat op waar dit op gebaseerd is."
    if print_prompt:
        print(prompt)
    response = client.chat.completions.create(
        model="gpt-4.1",
        messages=[
            {"role": "user", "content": prompt}], max_tokens=800
    )
    messages = [{'role': 'user', 'content': prompt}, {
        'role': 'assistant', 'content': response.choices[0].message.content}]
    # print("\t", response.choices[0].message.content)
    return messages, response.choices[0].message.content, response.usage.total_tokens

# debugg of start conversatie
def start_conversation(client, input_messages):
    messages = input_messages
    while True:
        user_input = input("User: \n\t")
        if user_input == 'q':
            break
        print("--------------------------------------")
        messages.append({'role': 'user', 'content': user_input})

        response = client.chat.completions.create(
            model="gpt-4.1",
            messages=messages, max_tokens=800
        )
        assistant_reply = response.choices[0].message.content
        print("Assistant:\n\n", assistant_reply)
        print("--------------------------------------")
        messages.append({"role": "assistant", "content": assistant_reply})
        
# wat post functies
def remove_short_words(match):
    content = match.group(1)
    filtered = ' '.join([w for w in content.split() if len(w) > 7])
    return f"({filtered})"

def get_score(correct_score, result_final):
    correct_score = ast.literal_eval(correct_score.item())
    score_compleet = 0
    score_2 = 0
    for item in result_final:
        break
        if item in correct_score:
            score_compleet += 1
            score_2 +=1
        if item not in correct_score:
            score_compleet -= 1  
    if len(correct_score) != 0:
        procent_extra_is_fout = score_compleet * 100 / len(correct_score)
        procent_extra_niet_fout = score_2 * 100 / len(correct_score)
    else:
        procent_extra_is_fout = 0
        procent_extra_niet_fout = 100
        
    if (len(correct_score) == 0) & (len(result_final) == 0):
        procent_extra_is_fout = 100
        procent_extra_niet_fout = 100
    return procent_extra_is_fout, procent_extra_niet_fout


# =============================================================================
# Parse de bestanden
# =============================================================================

# Loop door documenten en schrijf json vectorized
for model in ["text-embedding-ada-002", "text-embedding-3-large"]:
    for i, dossier in enumerate(os.listdir(data_directory)):


        # Prepare the data
        filepath = data_directory + dossier
        with pdfplumber.open(filepath) as pdf:
            # parse & vectorize

            tables = retrieve_table(client_4o, filepath)
            toc = TOC(pdf, find_start_page(pdf))
        
            json_var = vectorize_json(
                toc, dossier, i, pdf, tables, model, page_based=False)
            # Code to write to disk
            output = {}
            for item_x, item in enumerate(json_var):

                output[str(item_x) + '_' + json_var[item_x]["paragraph"]] = {
                    'id': json_var[item_x]['id'],
                    'dossier': dossier.replace('.pdf', ''),
                    'paragraph': json_var[item_x]['paragraph'],
                    'paragraph_number': json_var[item_x]['paragraph_number'],
                    'content': json_var[item_x]['content'],
                    'content_vector': json_var[item_x]['content_vector']
                }
            with open(quickStorage + f"/{model}_{dossier.replace('.pdf', '')}.json", "w") as f:
                json.dump(output, f, indent=4)

# =============================================================================
# Bevraag met vectoren de LLM
# =============================================================================

#Query input
data_format_filepath = "/data/data format.xlsx"
data_format = pd.ExcelFile(data_format_filepath)
sheet_names = data_format.sheet_names
prompts = pd.read_excel(data_format_filepath,
                        sheet_name='Variabelen prompts', engine='openpyxl')

start_time = time.time()

#Define output vars
result = pd.DataFrame()
info = pd.DataFrame()
usage = {}
score = pd.DataFrame(columns=["dossier", "variabele", "score_extra_isFout", "score_extra_nietFout", "run"])

for i, dossier in enumerate(os.listdir(quickStorage)):
    used_tokens = 0
    #choose vector with preferred model
    if embedding_model[model_n] not in dossier:
        continue
 
    # Open files
    with open(quickStorage + dossier) as file:
        json_var = json.load(file)
    
    toc = "\n".join(str(item) for item in json_var.keys())
    
    # Loop and ask model
    result_top_row = {'dossier': dossier.replace(
        embedding_model[model_n], '').replace('.pdf', '')}
    for index, row in prompts[16:18].iterrows():
       
        #Gebruik cosine similarity om 5 paragraven met een vergelijkbare semantische gelijknis te vinden
        text = cosine_similarity_func(json_var, embed_query(row['Prompt'] + str(row['tags']), embedding_model[model_n]), paragraven_n=3, print_score=False)
        
        #Voer meerdere runs uit en baseer eind antwoord op meest voorkomende antwoord
        if row['multiple_runs']:
            result_lijst = []
            for iteration in range(1, n+1):
                try:
                    #vraag het LLM, input = vraag + randvoorwaarden en paragraven uit cosine_similarity
                    message_hist, result_llm, tokens = ask_llm(client, text, row['Prompt'], row['randvoorwaarden_output'], toc,
                                                       print_prompt=False,
                                                       toelichting=False)
                except:
                    pass
                used_tokens += tokens
                try:
                    result_llm = re.sub(r"\(([^])]*)\)", remove_short_words, result_llm).replace(" ()", "")
                except:
                    pass
                result_lijst.append(result_llm)
                # start_conversation(client, message_hist)
            result_lijst = [x.replace('"', '').replace("'", "").replace("| ", "*8*") for x in result_lijst if pd.notna(x)]
            individuals = []
            for result_lijst_x in result_lijst:
                items = [item.strip().replace("\n", "").lower()
                         for item in result_lijst_x.split("*8*")]
                individuals.extend(items)

            counts = Counter(individuals)
            result_final = [stof for stof,count in counts.items() if count >= (0.27 * n)]
            #Filter vaak voorkomende fouten zoals opgeven verzamelnamen of alleen de concentratie
            stofgroepen = ["farmaceutische middelen", "medicijnresten", "overige verontreinigende stoffen", "veterinaire geneesmiddelen", "macro-parameters", "metalen"]
            result_final = [item for item in result_final if item not in stofgroepen]
            result_final = [s for s in result_final if any(c.isalpha() for c in s)]
            result_final.sort() 
            #print(result_final)
        else:
            message_hist, result_llm, tokens = ask_llm(
                client, text, row['Prompt'], row['randvoorwaarden_output'], toc, print_prompt=False, toelichting=False)
            used_tokens += tokens
            # print(result_llm)
            result_final = result_llm
            # start_conversation(client, message_hist)
        
        result_top_row[row['Variabele']] = result_final
        
        
    usage[dossier] = used_tokens
    result = pd.concat([result, pd.DataFrame([result_top_row])])
    
result = pd.merge(result, info, on='dossier')
result.to_excel(result_directory + 'behandelmethodes.xlsx', index=False)
info.to_excel(result_directory + 'all_dossiers_info.xlsx', index=False)
end_time = time.time()
duration = end_time-start_time

print(f"Duration: {duration:.4f} seconds")

def to_list(val):
    if isinstance(val, list):
        return val
    elif isinstance(val, str):
        try:
            return ast.literal_eval(val)
        except:
            return []
    else:
        return []

def split_stoffen(val):
    if pd.isna(val) or val == []:
        return []
    elif isinstance(val, list):
        return val
    elif isinstance(val, str):
        # Split op komma en strip eventuele spaties
        return [stof.strip() for stof in val.split(',')]
    else:
        return []








    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    