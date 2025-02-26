import nltk
import streamlit as st
import speech_recognition as sr
import pyttsx3
from nltk.tokenize import word_tokenize
from docx import Document
import pandas as pd
import threading

# Ensure NLTK data is downloaded
nltk.download('punkt')

# Initialize text-to-speech engine
engine = pyttsx3.init()
engine.setProperty('rate', 150)  # Adjust speaking speed

# File paths for images
mic_icon = r"C:\Users\User\Documents\ChatbotProject\Image\mic_button.png"  # Click to Speak
listening_icon = r"C:\Users\User\Documents\ChatbotProject\Image\listening.gif"  # Listening animation
bot_icon = r"C:\Users\User\Documents\ChatbotProject\Image\bot_icon.png"  # Bot avatar

# Function to convert text to speech
def speak(text):
    def run_speech():
        engine = pyttsx3.init()
        engine.say(text)
        engine.runAndWait()
    
    threading.Thread(target=run_speech, daemon=True).start()

# Function to read DOCX and extract text
def read_docx(file):
    doc = Document(file)
    text = [para.text for para in doc.paragraphs]
    return '\n'.join(text)

# Function to search for the phrase in the document
def search_paragraphs(doc_text, search_words):
    paragraphs = doc_text.split('\n')
    result = [para for para in paragraphs if all(word.lower() in para.lower() for word in search_words)]
    return result

# Function to handle greetings and conversations
def handle_greetings(user_input):
    greetings = ["hi", "hello", "hey", "good morning", "good evening", "good night"]
    user_input = user_input.lower()

    if any(greet in user_input for greet in greetings):
        return "Hello! How can I assist you today?"
    elif "how are you" in user_input:
        return "I am fine, thank you! How are you?"
    elif any(phrase in user_input for phrase in ["i am good", "i am fine", "i am good too"]):
        return "Good to know that! How may I help you?"
    elif any(phrase in user_input for phrase in ["i am not good", "i am not fine", "i am ok ok"]):
        return "What happened? Is there anything I can help with?"
    elif "where are you from" in user_input:
        return "I am from Trivandrum UST Campus A-Z Account, thank you! How may I assist you?"
    elif "can you get me some details from document" in user_input:
        return "For that, please enter at least 2 keywords from the document."
    
    return None

# Function to authenticate user from an Excel file
def authenticate_user(username, password, excel_file="users.xlsx"):
    try:
        df = pd.read_excel(excel_file)
        df.columns = df.columns.str.strip()
        df["Username"] = df["Username"].astype(str).str.strip()
        df["Password"] = df["Password"].astype(str).str.strip()
        user_match = (df["Username"] == username) & (df["Password"] == password)
        return user_match.any()
    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")
        return False

# Function to recognize voice input
def recognize_speech():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        #st.image(listening_icon, caption="🎤 Listening... Speak Now!", use_column_width=True)  # Listening GIF
        try:
            audio = recognizer.listen(source, timeout=5)
            return recognizer.recognize_google(audio)
        except sr.UnknownValueError:
            return "Sorry, I couldn't understand the audio."
        except sr.RequestError:
            return "Error with the speech recognition service."

# Streamlit UI
st.title("🤖 Finmantra Chatbot - Secure Login")

# User login
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.subheader("🔒 Login to Access the Chatbot")
    username = st.text_input("Username:", key="username")
    password = st.text_input("Password:", type="password", key="password")
    
    if st.button("Login"):
        if authenticate_user(username, password):
            st.session_state.authenticated = True
            st.success("✅ Login successful! Welcome to the chatbot.")
            st.experimental_rerun()
        else:
            st.error("❌ Invalid username or password.")
else:
    st.success("✅ You are logged in!")

    # Upload file once authenticated
    st.subheader("📂 Upload a .docx File")
    uploaded_file = st.file_uploader("Drag & drop your file here", type=["docx"], key="file_uploader")
    
    if uploaded_file:
        st.success(f"📄 File uploaded successfully: {uploaded_file.name}")

        # Read document content
        doc_text = read_docx(uploaded_file)

        # Chatbot interaction
        st.subheader("💬 Ask Finmantra Bot About the Document")

        # **Voice Input Button**
        st.markdown("### 🎙️ Use Voice Command")
        col1, col2, col3 = st.columns([1, 3, 1])
        with col2:
            if st.button("Click to Speak 🎤"):
               # st.image(mic_icon, caption="🎤 Speak Now", use_column_width=False, width=50)  # Mic Button
                user_input = recognize_speech()
                if user_input:
                    st.write(f"**You Said:** {user_input}")  

        # **Text Input Field**
        user_text_input = st.text_input("Or type your query here:", key="text_input")

        # **Process User Input (Either from Voice or Text)**
        final_input = user_text_input if user_text_input else user_input if 'user_input' in locals() else None

        if final_input:
            st.write(f"**You:** {final_input}")

            # **Process the Query**
            if final_input.lower() == "bye":
             #   st.image(bot_icon, width=50)
                st.write("🤖 Finmantra Bot: Goodbye! Have a great day! 😊")
                speak("Goodbye! Have a great day!")
            else:
                greeting_response = handle_greetings(final_input)
                if greeting_response:
                    #st.image(bot_icon, width=50)
                    st.write(f"🤖 FinMantra Bot : {greeting_response}")
                    speak(greeting_response)
                else:
                    search_words = final_input.split()
                    if len(search_words) > 1:
                        paragraphs = search_paragraphs(doc_text, search_words)
                        if paragraphs:
                            response_text = "Here's what I found:\n" + "\n".join(paragraphs[:3])
                        #    st.image(bot_icon, width=50)
                            st.write(f"🤖 FinMantra Bot: {response_text}")
                            speak(response_text)
                        else:
                         #   st.image(bot_icon, width=50)
                            st.write("🤖 FinMantra Bot: I couldn't find any matching paragraph.")
                            speak("I couldn't find any matching paragraph.")
                    else:
                       # st.image(bot_icon, width=50)
                        st.write("🤖 FinMantra Bot: Please provide at least two words to search for.")
                        speak("Please provide at least two words to search for.")