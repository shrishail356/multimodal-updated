import streamlit as st
from llm_chains import load_normal_chain, load_pdf_chat_chain
from langchain.memory import StreamlitChatMessageHistory
from streamlit_mic_recorder import mic_recorder
from utils import save_chat_history_json, get_timestamp, load_chat_history_json
from image_handler import handle_image
from audio_handler import transcribe_audio
from pdf_handler import add_documents_to_db
from html_templates import get_bot_template, get_user_template, css
from dotenv import load_dotenv
import yaml
import os
import time
from datetime import datetime
import json
from document_handler import DocumentProcessor
from utils import generate_file_preview, get_file_extension
import re
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Load environment variables and configuration
load_dotenv()
with open("config.yaml", "r") as f:
    config = yaml.safe_load(f)

# Create chat_sessions directory if it doesn't exist
os.makedirs(config["chat_history_path"], exist_ok=True)

AVAILABLE_MODELS = [
    # Production Models
    "gemma2-9b-it",  # Google
    "llama-3.3-70b-versatile",  # Meta
    "llama-3.1-8b-instant",  # Meta
    "llama-guard-3-8b",  # Meta
    "llama3-70b-8192",  # Meta
    "llama3-8b-8192",  # Meta
    
    # Preview Models
    "allam-2-7b",  # SDAIA
    "deepseek-r1-distill-llama-70b",  # DeepSeek
    "meta-llama/llama-4-maverick-17b-128e-instruct",  # Meta
    "meta-llama/llama-4-scout-17b-16e-instruct",  # Meta
    "mistral-saba-24b",  # Mistral
    "qwen-qwq-32b"  # Alibaba Cloud
]

def load_chain(chat_history):
    model_name = st.session_state.selected_model
    if st.session_state.pdf_chat:
        print("loading pdf chat chain")
        return load_pdf_chat_chain(chat_history, model_name=model_name)
    return load_normal_chain(chat_history, model_name=model_name)

def clear_input_field():
    if st.session_state.user_question == "":
        st.session_state.user_question = st.session_state.user_input
        st.session_state.user_input = ""

def set_send_input():
    st.session_state.send_input = True
    clear_input_field()

def toggle_pdf_chat():
    st.session_state.pdf_chat = True

def save_chat_history():
    if st.session_state.history != []:
        if st.session_state.session_key == "new_session":
            st.session_state.new_session_key = get_timestamp() + ".json"
            save_chat_history_json(st.session_state.history, config["chat_history_path"] + st.session_state.new_session_key)
        else:
            save_chat_history_json(st.session_state.history, config["chat_history_path"] + st.session_state.session_key)

def display_chat_history(messages, thinking_steps_history):
    """Display complete chat history with command execution feedback."""
    for i, message in enumerate(messages):
        timestamp = datetime.fromtimestamp(message.additional_kwargs.get('timestamp', time.time())).strftime('%Y-%m-%d %H:%M:%S')
        
        if message.type == "human":
            st.write(get_user_template(message.content, timestamp), unsafe_allow_html=True)
        else:
            # Check if this is a command execution response
            content = message.content
            if "Successfully" in content or "Error" in content:
                content = f"🔧 Command Execution: {content}"
            
            steps = thinking_steps_history.get(i, None)
            st.write(get_bot_template(content, timestamp, steps), unsafe_allow_html=True)

def add_customization_settings():
    st.sidebar.title("Interface Settings")
    
    # Theme Selection
    if 'theme_settings' not in st.session_state:
        st.session_state.theme_settings = {
            'primary_bg': '#1a1c24',
            'message_user_bg': '#2b313e',
            'message_bot_bg': '#3a4152',
            'accent_color': '#4CAF50',
            'text_color': '#ffffff',
            'text_secondary': '#a0aec0',
            'font_size': '16',
            'message_spacing': '1.5',
            'avatar_size': '50'
        }
    
    with st.sidebar.expander("Theme Settings"):
        # Color settings
        st.session_state.theme_settings['primary_bg'] = st.color_picker(
            "Background Color", 
            st.session_state.theme_settings['primary_bg']
        )
        st.session_state.theme_settings['message_user_bg'] = st.color_picker(
            "User Message Color", 
            st.session_state.theme_settings['message_user_bg']
        )
        st.session_state.theme_settings['message_bot_bg'] = st.color_picker(
            "Bot Message Color", 
            st.session_state.theme_settings['message_bot_bg']
        )
        st.session_state.theme_settings['accent_color'] = st.color_picker(
            "Accent Color", 
            st.session_state.theme_settings['accent_color']
        )
        st.session_state.theme_settings['text_color'] = st.color_picker(
            "Text Color", 
            st.session_state.theme_settings['text_color']
        )
        
        # Layout settings
        st.session_state.theme_settings['font_size'] = st.slider(
            "Font Size (px)", 
            12, 24, 
            int(st.session_state.theme_settings['font_size'])
        )
        st.session_state.theme_settings['message_spacing'] = st.slider(
            "Message Spacing (rem)", 
            0.5, 3.0, 
            float(st.session_state.theme_settings['message_spacing'])
        )
        st.session_state.theme_settings['avatar_size'] = st.slider(
            "Avatar Size (px)", 
            30, 80, 
            int(st.session_state.theme_settings['avatar_size'])
        )

    # Save/Load Theme
    with st.sidebar.expander("Save/Load Theme"):
        # Save current theme
        if st.button("Save Current Theme"):
            with open("saved_theme.json", "w") as f:
                json.dump(st.session_state.theme_settings, f)
            st.success("Theme saved!")
            
        # Load saved theme
        if os.path.exists("saved_theme.json"):
            if st.button("Load Saved Theme"):
                with open("saved_theme.json", "r") as f:
                    st.session_state.theme_settings = json.load(f)
                st.success("Theme loaded!")

def apply_custom_theme():
    custom_css = f"""
    <style>
    :root {{
        --primary-bg: {st.session_state.theme_settings['primary_bg']};
        --message-user-bg: {st.session_state.theme_settings['message_user_bg']};
        --message-bot-bg: {st.session_state.theme_settings['message_bot_bg']};
        --accent-color: {st.session_state.theme_settings['accent_color']};
        --text-color: {st.session_state.theme_settings['text_color']};
        --text-secondary: {st.session_state.theme_settings['text_secondary']};
    }}

    .chat-message {{
        font-size: {st.session_state.theme_settings['font_size']}px !important;
        margin-bottom: {st.session_state.theme_settings['message_spacing']}rem !important;
    }}

    .chat-message .avatar img {{
        width: {st.session_state.theme_settings['avatar_size']}px !important;
        height: {st.session_state.theme_settings['avatar_size']}px !important;
    }}
    </style>
    """
    st.markdown(custom_css, unsafe_allow_html=True)

def main():
    st.set_page_config(page_title="Multimodal Chat with Groq", page_icon="🤖", layout="wide")
    st.title("Multimodal Chat App with Groq Integration")
    st.write(css, unsafe_allow_html=True)
    
    # Initialize session states
    if "send_input" not in st.session_state:
        st.session_state.session_key = "new_session"
        st.session_state.send_input = False
        st.session_state.user_question = ""
        st.session_state.new_session_key = None
        st.session_state.session_index_tracker = "new_session"
        st.session_state.thinking_steps_history = {}
        st.session_state.current_thinking_steps = []
        st.session_state.message_timestamps = {}  # New: store timestamps separately
        st.session_state.pdf_chat = False

    # Initialize thinking steps history
    if "thinking_steps_history" not in st.session_state:
        st.session_state.thinking_steps_history = {}

    # Groq status indicator
    with st.sidebar:
        st.write("🤖 Using Groq API")
        if not os.getenv("GROQ_API_KEY"):
            st.error("⚠️ GROQ_API_KEY not found in environment variables")
        
        # Initialize selected_model in session state if not present
        if "selected_model" not in st.session_state:
            st.session_state.selected_model = "llama-3.3-70b-versatile"
        
        st.session_state.selected_model = st.selectbox(
            "Select Model",
            options=AVAILABLE_MODELS,
            index=AVAILABLE_MODELS.index(st.session_state.selected_model)
        )
    
    # Chat session management
    st.sidebar.title("Chat Sessions")
    chat_sessions = ["new_session"] + sorted(os.listdir(config["chat_history_path"]))
    
    if st.session_state.session_key == "new_session" and st.session_state.new_session_key is not None:
        st.session_state.session_index_tracker = st.session_state.new_session_key
        st.session_state.new_session_key = None

    index = chat_sessions.index(st.session_state.session_index_tracker)
    selected_session = st.sidebar.selectbox("Select a chat session", chat_sessions, key="session_key", index=index)
    
    # PDF chat toggle
    st.sidebar.toggle("PDF Chat", key="pdf_chat", value=False)

    # Load chat history
    if selected_session != "new_session":
        st.session_state.history = load_chat_history_json(config["chat_history_path"] + selected_session)
    else:
        st.session_state.history = []

    chat_history = StreamlitChatMessageHistory(key="history")
    
    # Main layout
    chat_container = st.container()
    thinking_container = st.container()
    input_container = st.container()

    # File upload and preview section in sidebar
    with st.sidebar:
        st.title("Upload Files")
        
        # Audio upload
        uploaded_audio = st.file_uploader("Upload an audio file", type=["wav", "mp3", "ogg"])
        
        # Image upload
        uploaded_image = st.file_uploader("Upload an image file", type=["jpg", "jpeg", "png"])
        
        # Document upload
        uploaded_documents = st.file_uploader(
            "Upload documents",
            accept_multiple_files=True,
            type=["pdf", "txt", "csv", "xlsx", "xls", "docx"],
            key="document_upload"
        )
        
        # File Preview Section
        if uploaded_documents:
            st.write("### File Previews")
            for doc in uploaded_documents:
                with st.expander(f"Preview: {doc.name}"):
                    file_type = get_file_extension(doc.name)
                    preview = generate_file_preview(doc, file_type)
                    
                    if file_type in ['png', 'jpg', 'jpeg']:
                        st.image(doc)
                    elif file_type == 'pdf':
                        try:
                            import pdfplumber
                            with pdfplumber.open(doc) as pdf:
                                st.write(pdf.pages[0].extract_text()[:1000] + '...')
                        except Exception as e:
                            st.error(f"Error previewing PDF: {str(e)}")
                    else:
                        st.write(preview)
                    
                    # Process button for each file
                    if st.button(f"Process {doc.name}", key=f"process_{doc.name}"):
                        with st.spinner(f"Processing {doc.name}..."):
                            try:
                                doc_processor = DocumentProcessor()
                                content = doc_processor.process_file(doc, file_type)
                                
                                # Handle the processed content
                                if file_type == "pdf":
                                    add_documents_to_db([doc])
                                else:
                                    # For other document types, create a conversation entry
                                    llm_chain = load_chain(chat_history)
                                    response = llm_chain.run(f"Summarize this document: {content[:2000]}")
                                    
                                    current_time = time.time()
                                    message_index = len(chat_history.messages)
                                    st.session_state.message_timestamps[message_index] = current_time
                                    chat_history.add_user_message(f"Processed document: {doc.name}")
                                    
                                    st.session_state.message_timestamps[message_index + 1] = current_time
                                    chat_history.add_ai_message(response)
                                
                                st.success(f"Successfully processed {doc.name}")
                            except Exception as e:
                                st.error(f"Error processing {doc.name}: {str(e)}")

    # Process uploaded files
    pdf_files = []
    if uploaded_documents:
        pdf_files = [doc for doc in uploaded_documents if get_file_extension(doc.name) == 'pdf']
        if pdf_files:
            with st.spinner("Processing PDF files..."):
                add_documents_to_db(pdf_files)

    if uploaded_audio:
        with st.spinner("Transcribing audio..."):
            transcribed_audio = transcribe_audio(uploaded_audio.getvalue())
            llm_chain = load_chain(chat_history)
            response = llm_chain.run("Summarize this text: " + transcribed_audio)
            current_time = time.time()
            
            # Store timestamp before adding message
            message_index = len(chat_history.messages)
            st.session_state.message_timestamps[message_index] = current_time
            chat_history.add_user_message("Audio transcription: " + transcribed_audio)
            
            st.session_state.message_timestamps[message_index + 1] = current_time
            chat_history.add_ai_message(response)

    # Handle uploaded image before the input section
    if uploaded_image is not None:
        with st.spinner("Processing image..."):
            image_bytes = uploaded_image.getvalue()
            
            # Get user's query about the image
            image_query = st.text_input(
                "What would you like to know about this image?",
                value="Describe this image in detail.",
                key="image_query"
            )
            
            if st.button("Analyze Image"):
                success, image_response = handle_image(
                    image_bytes=image_bytes,
                    user_message=image_query
                )
                
                current_time = time.time()
                message_index = len(chat_history.messages)
                
                # Add the image query to chat history
                st.session_state.message_timestamps[message_index] = current_time
                chat_history.add_user_message(f"[Image Analysis Request]: {image_query}")
                
                # Add the analysis response to chat history
                st.session_state.message_timestamps[message_index + 1] = current_time
                if success:
                    chat_history.add_ai_message(image_response)
                else:
                    chat_history.add_ai_message(f"⚠️ {image_response}")
                
                # Display the image with the response
                st.image(uploaded_image, caption="Uploaded Image", use_column_width=True)

    # Input section
    with input_container:
        user_input = st.text_input("Type your message here", key="user_input", on_change=set_send_input)
        col1, col2 = st.columns(2)
        with col1:
            voice_recording = mic_recorder(
                start_prompt="Start recording",
                stop_prompt="Stop recording",
                just_once=True
            )
        with col2:
            send_button = st.button("Send", key="send_button", on_click=clear_input_field)

    # Handle voice recording
    if voice_recording:
        with st.spinner("Processing voice recording..."):
            transcribed_audio = transcribe_audio(voice_recording["bytes"])
            llm_chain = load_chain(chat_history)
            response = llm_chain.run(transcribed_audio)
            current_time = time.time()
            
            message_index = len(chat_history.messages)
            st.session_state.message_timestamps[message_index] = current_time
            chat_history.add_user_message("Voice input: " + transcribed_audio)
            
            st.session_state.message_timestamps[message_index + 1] = current_time
            chat_history.add_ai_message(response)

    # Process user input and generate response
    if send_button or st.session_state.send_input:
        current_time = time.time()
        
        if st.session_state.user_question:
            user_message = st.session_state.user_question
            llm_chain = load_chain(chat_history)
            
            # Check if input is a computer command
            is_command, cmd_type, command_chain = llm_chain.llm.interpret_command(user_message)
            
            if is_command:
                message_index = len(chat_history.messages)
                st.session_state.message_timestamps[message_index] = current_time
                chat_history.add_user_message(user_message)
                
                response = llm_chain.llm.execute_command_chain(command_chain)
                
                st.session_state.message_timestamps[message_index + 1] = current_time
                chat_history.add_ai_message(f"Executed command chain: {response}")
                st.session_state.user_question = ""
            
            else:
                # Handle normal conversation
                message_index = len(chat_history.messages)
                st.session_state.message_timestamps[message_index] = current_time
                chat_history.add_user_message(user_message)
                
                # Process and display thinking steps
                with thinking_container:
                    thinking_status = st.empty()
                    response_placeholder = st.empty()
                    
                    current_steps = []
                    for step_num, (title, content, thinking_time) in enumerate(
                        llm_chain.run_with_steps(user_message), 1
                    ):
                        if "Final Answer" not in title:
                            current_steps.append((title, content))
                            thinking_status.markdown(f"🤔 Step {step_num}: {title}")
                            with response_placeholder.container():
                                st.markdown(f"**{title}**")
                                st.markdown(content)
                                st.markdown(f"*Thinking time: {thinking_time:.2f}s*")
                        else:
                            # Store final answer and thinking steps
                            message_index = len(chat_history.messages)
                            st.session_state.message_timestamps[message_index] = time.time()
                            chat_history.add_ai_message(content)
                            st.session_state.thinking_steps_history[message_index] = current_steps
            
            st.session_state.user_question = ""
        
        st.session_state.send_input = False

    # Modified display_chat_history function
    def display_chat_history(messages, thinking_steps_history):
        """Display complete chat history with command execution feedback."""
        for i, message in enumerate(messages):
            timestamp = datetime.fromtimestamp(
                st.session_state.message_timestamps.get(i, time.time())
            ).strftime('%Y-%m-%d %H:%M:%S')
            
            if message.type == "human":
                st.write(get_user_template(message.content, timestamp), unsafe_allow_html=True)
            else:
                steps = thinking_steps_history.get(i, None)
                st.write(get_bot_template(message.content, timestamp, steps), unsafe_allow_html=True)

    # Display chat history
    with chat_container:
        st.markdown("### Chat History")
        if chat_history.messages:
            display_chat_history(chat_history.messages, st.session_state.thinking_steps_history)
        else:
            st.info("Start a conversation by typing a message or uploading a file.")

    # Add scroll to top button
    st.markdown('''
        <button onclick="window.scrollTo(0,0)" class="scroll-top-btn">↑</button>
    ''', unsafe_allow_html=True)

    # Save chat history
    save_chat_history()

    # Add customization settings in sidebar
    add_customization_settings()
    
    # Apply custom theme
    apply_custom_theme()

if __name__ == "__main__":
    main()