import streamlit as st
import msal
import requests
import base64
import time
import re
import io
from docx import Document
from PyPDF2 import PdfReader
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from datetime import datetime, timedelta
import streamlit_pagination as stp
import zipfile
 
# Azure AD app details
client_id = st.secrets["CLIENT_ID"]
client_secret = st.secrets["CLIENT_SECRET"]
tenant_id = st.secrets["TENANT_ID"]
authority_url = f"https://login.microsoftonline.com/{tenant_id}"
redirect_uri = st.secrets["URL"]
 
# Define the scopes required for accessing SharePoint
# scopes=["Files.ReadWrite.All", "Sites.Read.All", "User.Read"]
 
# MSAL configuration
app = msal.ConfidentialClientApplication(
    client_id,
    authority=authority_url,
    client_credential=client_secret,
)
 
# Streamlit UI
st.title("📂 SharePoint File Downloader and Query Chatbot")
 
# Authentication flow
def get_auth_url():
    auth_url = app.get_authorization_request_url(
        scopes=["Files.ReadWrite.All", "Sites.Read.All", "User.Read"], redirect_uri=redirect_uri)
    return auth_url
 
def get_token_from_code(auth_code):
    result = app.acquire_token_by_authorization_code(
        auth_code, scopes=["Files.ReadWrite.All", "Sites.Read.All", "User.Read"], redirect_uri=redirect_uri)
    return result
 
# Cache the authentication headers
@st.cache_resource
def get_auth_headers(auth_code=None):
    if auth_code:
        token_response = get_token_from_code(auth_code)
        if 'access_token' in token_response:
            return {'Authorization': f'Bearer {token_response["access_token"]}'}
    return None
 
# Add this to initialize folder path in session state
if 'current_folder_path' not in st.session_state:
    st.session_state.current_folder_path = ""
 
# Update the list_items function to handle folder navigation and check for empty folders/sites
def list_items(url, headers, path="", prefix="", indent=""):
    items_list = []
    file_count = 0
    folder_count = 0
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        items = response.json().get('value', [])
        if not items:
            # Return a special item to indicate empty folder/site
            return [("empty", "", "This folder/site is empty", "Empty", "")], 0, 0
        for idx, item in enumerate(items, start=1):
            item_type = 'File' if 'file' in item else 'Folder'
            if item_type == 'File':
                file_count += 1
            else:
                folder_count += 1
            full_path = f"{path}/{item['name']}".lstrip("/")
            item_prefix = f"{prefix}{idx}"
            items_list.append(
                (item_prefix, full_path, f"{indent}{item_prefix}. {item['name']} ({item_type})", item_type, item['id']))
            if item_type == 'Folder' and st.session_state.current_folder_path == full_path:
                child_url = f"https://graph.microsoft.com/v1.0/sites/{site_info['id']}/drive/items/{item['id']}/children"
                sub_items_list, sub_file_count, sub_folder_count = list_items(
                    child_url, headers, full_path, f"{item_prefix}.", indent + "  ")
                items_list.extend(sub_items_list)
                file_count += sub_file_count
                folder_count += sub_folder_count
    return items_list, file_count, folder_count
 
def download_file(site_id, file_path, headers):
    file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/content"
    download_response = requests.get(file_url, headers=headers)
 
    if download_response.status_code == 200:
        return download_response.content, file_path.split("/")[-1]
    return None, None
 
def search_files(site_id, query, headers):
    search_url = f"https://graph.microsoft.com/v1.0/sites/{site_info['id']}/drive/root/search(q='{query}')"
    response = requests.get(search_url, headers=headers)
    if response.status_code == 200:
        items = response.json().get('value', [])
        return items
    return []
 
def read_file_content(file_content, file_name):
    content = ""
    try:
        if file_name.endswith('.txt'):
            content = file_content.decode('utf-8')
        elif file_name.endswith('.docx'):
            doc = Document(io.BytesIO(file_content))
            content = ' '.join([para.text for para in doc.paragraphs])
        elif file_name.endswith('.pdf'):
            pdf = PdfReader(io.BytesIO(file_content))
            content = ' '.join([page.extract_text() for page in pdf.pages])
        elif file_name.endswith('.csv'):
            df = pd.read_csv(io.StringIO(file_content.decode('utf-8')))
            content = df.to_string()
    except Exception as e:
        st.error(f"Error reading {file_name}: {str(e)}")
    return content
 
def preprocess_content(content):
    lines = content.split('\n')
    processed_lines = [line for line in lines if len(
        line.split()) > 3 and not line.strip().startswith('http')]
    return ' '.join(processed_lines)
 
def search_answer(question, file_contents):
    if not file_contents:
        return "I couldn't find any content in the files to answer the question."
 
    # Combine all file contents, keeping track of which document each sentence comes from
    all_content = []
    for doc_name, content in file_contents.items():
        sentences = [sent.strip()
                     for sent in content.split('.') if sent.strip()]
        all_content.extend([(sent, doc_name) for sent in sentences])
 
    # Separate sentences and document names
    sentences, doc_names = zip(*all_content)
    sentences = list(sentences)
 
    # Add the question to the sentences
    sentences.append(question)
 
    # Create TF-IDF vectors
    vectorizer = TfidfVectorizer()
    tfidf_matrix = vectorizer.fit_transform(sentences)
 
    # Compute cosine similarity
    cosine_similarities = cosine_similarity(
        tfidf_matrix[-1:], tfidf_matrix[:-1])[0]
 
    # Get indices of sentences with a similarity above the threshold
    relevant_sentence_indices = [i for i, score in enumerate(
        cosine_similarities) if score > 0.2]
 
    answers_dict = {}
 
    for idx in relevant_sentence_indices:
        sentence = sentences[idx].strip()
        # Remove citations and irrelevant information
        sentence = re.sub(r'\[[^\]]*\]', '', sentence)
        sentence = re.sub(r'\s+', ' ', sentence).strip()
 
        # Check if the sentence is complete and relevant
        if len(sentence.split()) > 5 and not sentence.startswith("Artificial intelligence") and "founding fathers of AI" not in sentence:
            doc_name = doc_names[idx]
            if doc_name not in answers_dict:
                answers_dict[doc_name] = []
            answers_dict[doc_name].append(sentence)
 
    if answers_dict:
        answer = "Here's what I found about your question:\n\n"
        for doc_name, relevant_sentences in answers_dict.items():
            combined_answer = ' '.join(relevant_sentences)
            combined_answer = combined_answer.capitalize()
            if not combined_answer.endswith('.'):
                combined_answer += '.'
            answer += f"**Source: {doc_name}**\n{combined_answer}\n\n"
    else:
        answer = "I'm sorry, but I couldn't find any relevant information to answer your question."
 
    return answer
 
# Function to download all files in a folder and create a zip archive
def download_folder_as_zip(site_id, folder_path, headers):
    folder_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder_path}:/children"
    response = requests.get(folder_url, headers=headers)
    if response.status_code == 200:
        items = response.json().get('value', [])
        folder_name = folder_path.split("/")[-1]
        zip_filename = f"{folder_name}.zip"
 
        with io.BytesIO() as zip_buffer:
            with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                for item in items:
                    if 'file' in item:  # Only include files, not subfolders
                        file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item['id']}/content"
                        file_response = requests.get(file_url, headers=headers)
                        if file_response.status_code == 200:
                            zip_file.writestr(
                                item['name'], file_response.content)
            zip_buffer.seek(0)
            return zip_buffer.read(), zip_filename
    return None, None
 
# New function to list accessible sites
def list_accessible_sites(headers):
    sites_url = "https://graph.microsoft.com/v1.0/sites?search=*"
    response = requests.get(sites_url, headers=headers)
    if response.status_code == 200:
        sites = response.json().get('value', [])
        return [(site['name'], site['webUrl']) for site in sites]
    return []
 
# Function to add a message to the chat history
def add_message(role, content):
    timestamp = datetime.now()
    st.session_state.messages.append({"role": role, "content": content, "timestamp": timestamp})
    with st.chat_message(role):
        st.markdown(content)
 
# Updated function to display chat history
def display_chat_history():
    st.sidebar.title("Chat History")
    st.sidebar.markdown("""
    <style>
    .chat-message {
        padding: 10px;
        border-radius: 10px;
        margin-bottom: 10px;
        font-size: 14px;
    }
    .user-message {
        background-color: #e6f3ff;
        border-left: 5px solid #2196F3;
    }
    .assistant-message {
        background-color: #f0f0f0;
        border-left: 5px solid #4CAF50;
    }
    .timestamp {
        font-size: 10px;
        color: #888;
        margin-top: 5px;
    }
    </style>
    """, unsafe_allow_html=True)
 
    current_time = datetime.now()
    for msg in reversed(st.session_state.messages):
        if current_time - msg["timestamp"] < timedelta(hours=24):
            message_class = "user-message" if msg['role'] == "user" else "assistant-message"
            st.sidebar.markdown(f"""
            <div class="chat-message {message_class}">
                <strong>{msg['role'].capitalize()}:</strong> {msg['content'][:100]}...
                <div class="timestamp">{msg['timestamp'].strftime('%Y-%m-%d %H:%M:%S')}</div>
            </div>
            """, unsafe_allow_html=True)
            # Main conversation flow starts here...
if 'auth_code' not in st.session_state:
    if st.button("Authenticate to use the app"):
        auth_url = get_auth_url()
        st.query_params["auth_url"] = auth_url
        st.rerun()
else:
    headers = get_auth_headers(st.session_state['auth_code'])
    if headers:
        if 'messages' not in st.session_state:
            st.session_state.messages = []
            st.session_state.items_dict = {}
            st.session_state.search_results_dict = {}
            st.session_state.file_contents = {}
 
        # Display chat history
        display_chat_history()
 
        if not st.session_state.messages:
            add_message("assistant", "Hello! I'm your SharePoint File Downloader and Query Assistant. How can I help you today?")
            add_message("assistant", "Would you like to view the sites you have access to? (Yes/No)")
 
        # Get user input for viewing sites
        user_input = st.text_input("Your response:")
        if user_input.lower() == 'yes':
            accessible_sites = list_accessible_sites(headers)
            if accessible_sites:
                sites_list = "\n".join([f"{idx + 1}. {site[0]} - {site[1]}" for idx, site in enumerate(accessible_sites)])
                add_message("assistant", f"Here are the SharePoint sites you have access to:\n\n{sites_list}")
                add_message("assistant", "Please enter the name of the SharePoint site you want to access (e.g., 'Chatbot_resource').")
            else:
                add_message("assistant", "I'm sorry, I couldn't retrieve the list of accessible sites. Please enter the name of the SharePoint site you want to access.")
        elif user_input.lower() == 'no':
            add_message("assistant", "Alright. Please enter the name of the SharePoint site you want to access (e.g., 'Chatbot_resource').")
 
        # Get site name input from user
        site_name = st.text_input("SharePoint Site Name:")
 
        # Define items per page
        ITEMS_PER_PAGE = 20
 
        # Function to get paginated items
        def get_paginated_items(items_list, page, items_per_page=ITEMS_PER_PAGE):
            start_index = (page - 1) * items_per_page
            end_index = start_index + items_per_page
            return items_list[start_index:end_index], len(items_list) // items_per_page + (1 if len(items_list) % items_per_page != 0 else 0)
 
        # Track the current page in the session state
        if 'current_page' not in st.session_state:
            st.session_state.current_page = 1
 
        # Add a function to display items with pagination and folder navigation
        def show_paginated_items(items_list, file_count, folder_count):
            current_items, total_pages = get_paginated_items(
                items_list, st.session_state.current_page)
 
            st.write(f"Total Files: {file_count}, Total Folders: {folder_count}")
           
            html_string = "<br>".join(
                f"<span style='font-size:15px'>{item[2]}</span>" for item in current_items)
            st.markdown(html_string, unsafe_allow_html=True)
 
            # Navigation buttons
            col1, col2, col3 = st.columns(3)
            with col1:
                if st.session_state.current_page > 1:
                    if st.button("Previous"):
                        st.session_state.current_page -= 1
                        st.rerun()
 
            with col2:
                st.write(
                    f"Page {st.session_state.current_page} of {total_pages}")
 
            with col3:
                if st.session_state.current_page < total_pages:
                    if st.button("Next"):
                        st.session_state.current_page += 1
                        st.rerun()
 
            # Add a back button to go back to the parent folder
            if st.session_state.current_folder_path:
                if st.button("Back"):
                    st.session_state.current_folder_path = "/".join(
                        st.session_state.current_folder_path.split("/")[:-1])
                    st.session_state.current_page = 1
                    st.rerun()
           
        # Update the main conversation flow to handle folder navigation and empty folders/sites
        if site_name:
            site_info_url = f'https://graph.microsoft.com/v1.0/sites/novintix.sharepoint.com:/sites/{site_name}'
            site_response = requests.get(site_info_url, headers=headers)
 
            if site_response.status_code == 200:
                site_info = site_response.json()
                if st.session_state.current_folder_path:
                    drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_info['id']}/drive/root:/{st.session_state.current_folder_path}:/children"
                else:
                    drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_info['id']}/drive/root/children"
 
                items_list, file_count, folder_count = list_items(
                    drive_url, headers, st.session_state.current_folder_path)
 
                if items_list and items_list[0][0] == "empty":
                    add_message(
                        "assistant", f"The SharePoint site '{site_name}' or the current folder is empty.")
                else:
                    st.session_state.items_dict = {
                        item[0]: (item[1], item[3]) for item in items_list}
                    add_message(
                        "assistant", f"Here are the contents of the SharePoint site '{site_name}':")
                    show_paginated_items(items_list, file_count, folder_count)  # Show items with pagination and counts
                    add_message(
                        "assistant", "Please provide the item number to view a folder's contents, to download('folder 1') or download a file. Or, ask a question by starting with 'Question:'")
            else:
                add_message(
                    "assistant", "I'm sorry, I couldn't access the SharePoint site. Please check the site name and try again.")
 
        # Main conversation flow handling folder download requests and questions
        if prompt := st.chat_input("You:"):
            add_message("user", prompt)
 
            if prompt.lower().startswith("question:"):
                question = prompt[9:].strip()  # Remove "Question:" prefix
                add_message("assistant", f"Searching for an answer to: '{question}'")
               
                # Search for relevant files
                search_results = search_files(site_info['id'], question, headers)
               
                if search_results:
                    # Download and read content of relevant files
                    for item in search_results:
                        file_content, file_name = download_file(site_info['id'], item['name'], headers)
                        if file_content and file_name:
                            content = read_file_content(file_content, file_name)
                            st.session_state.file_contents[file_name] = content
                   
                    # Search for answer in the downloaded content
                    answer = search_answer(question, st.session_state.file_contents)
                    add_message("assistant", answer)
                else:
                    add_message("assistant", "I'm sorry, I couldn't find any relevant files to answer your question.")
 
            elif prompt.lower().startswith("folder"):
                try:
                    folder_num = prompt.split()[1]
 
                    # Check if folder_num exists in items_dict
                    if folder_num in st.session_state.items_dict:
                        item_path, item_type = st.session_state.items_dict.get(
                            folder_num, (None, None))
 
                        if item_path and item_type == 'Folder':
                            # Download the entire folder as a zip
                            zip_content, zip_filename = download_folder_as_zip(
                                site_info['id'], item_path, headers)
                            if zip_content and zip_filename:
                                add_message(
                                    "assistant", f"Great! I've successfully downloaded the folder '{zip_filename}' for you.")
                                b64 = base64.b64encode(zip_content).decode()
                                href = f'<a href="data:application/zip;base64,{b64}" download="{zip_filename}">Click here to download {zip_filename}</a>'
                                st.markdown(href, unsafe_allow_html=True)
                                add_message(
                                    "assistant", "Is there anything else you'd like to do? (Yes/No)")
                            else:
                                add_message(
                                    "assistant", "I'm sorry, I couldn't download the folder. It might be empty or there was an error.")
                        else:
                            add_message(
                                "assistant", "The specified item is not a folder. Please check the folder number and try again.")
                    else:
                        add_message(
                            "assistant", f"No folder found with the number {folder_num}. Please check the number and try again.")
                except IndexError:
                    add_message(
                        "assistant", "Please specify the folder number after 'folder'. For example: 'folder 2'.")
 
            elif prompt.isdigit() or (prompt.replace('.', '').isdigit() and prompt.count('.') == 1):
                item_path, item_type = st.session_state.items_dict.get(
                    prompt, (None, None))
                if item_path:
                    if item_type == 'Folder':
                        # Navigate into the folder
                        st.session_state.current_folder_path = item_path
                        st.session_state.current_page = 1
                        st.rerun()
                    elif item_type == 'File':
                        # Download the file
                        file_content, file_name = download_file(
                            site_info['id'], item_path, headers)
                        if file_content and file_name:
                            add_message(
                                "assistant", f"Great! I've successfully downloaded '{file_name}' for you.")
                            b64 = base64.b64encode(file_content).decode()
                            href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">Click here to download {file_name}</a>'
                            st.markdown(href, unsafe_allow_html=True)
                            add_message(
                                "assistant", "Is there anything else you'd like to do? (Yes/No)")
                else:
                    add_message(
                        "assistant", "I'm sorry, the item number you provided does not exist. Please check the item number and try again.")
 
            elif "no" in prompt.lower():
                add_message(
                    "assistant", "Alright! It was great helping you. Have a wonderful day!")
                time.sleep(2)
                st.rerun()
 
            elif "yes" in prompt.lower():
                add_message(
                    "assistant", "Please enter your query to search for related files or ask a question starting with 'Question:'")
 
            else:
                # Handle search queries
                query = prompt
                add_message(
                    "assistant", f"Searching for files related to '{query}'...")
                search_results = search_files(site_info['id'], query, headers)
 
                if search_results:
                    search_results_list = [
                        f"{idx + 1}. {item['name']} (File)" for idx, item in enumerate(search_results)]
                    search_results_dict = {
                        str(idx + 1): item['id'] for idx, item in enumerate(search_results)}
 
                    st.session_state.search_results_dict = search_results_dict
                    search_results_text = "\n".join(search_results_list)
                    add_message(
                        "assistant", f"Here are the files related to your query:\n\n{search_results_text}")
                    add_message(
                        "assistant", "Please provide the file number to download.")
 
                else:
                    add_message(
                        "assistant", "No files found related to your query. Please try again with a different query.")
 
                if st.session_state.get('search_results_dict') and prompt in st.session_state['search_results_dict']:
                    file_id = st.session_state['search_results_dict'][prompt]
                    file_info_url = f"https://graph.microsoft.com/v1.0/sites/{site_info['id']}/drive/items/{file_id}"
                    file_info_response = requests.get(
                        file_info_url, headers=headers)
 
                    if file_info_response.status_code == 200:
                        file_info = file_info_response.json()
                        file_name = file_info['name']
                        download_url = file_info['@microsoft.graph.downloadUrl']
                        file_content = requests.get(download_url).content
                        add_message(
                            "assistant", f"Great! I've successfully downloaded '{file_name}' for you.")
 
                        # Create a download button
                        b64 = base64.b64encode(file_content).decode()
                        href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">Click here to download {file_name}</a>'
                        st.markdown(href, unsafe_allow_html=True)
 
                        add_message(
                            "assistant", "Is there anything else you'd like to do? (Yes/No)")
                    else:
                        add_message(
                            "assistant", "I'm sorry, I couldn't retrieve the file information. Please try again.")
 
        # Add a button to clear the conversation
        if st.button("Clear Conversation"):
            st.session_state.messages = []
            st.rerun()
 
# To handle the redirection and capture the auth code
if 'auth_code' not in st.session_state and 'code' in st.query_params:
    st.session_state.auth_code = st.query_params['code']
    st.rerun()
 
# Redirect to authentication URL
if 'auth_url' in st.query_params:
    st.markdown(
        f'<meta http-equiv="refresh" content="0; url={st.query_params["auth_url"]}">', unsafe_allow_html=True)
