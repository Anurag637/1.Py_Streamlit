import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
from win32com.client import Dispatch
import pythoncom

# Set page configuration - Must be the first Streamlit command
st.set_page_config(
    page_title="Email Spam Classifier",
    layout="centered",
    initial_sidebar_state="expanded",
    page_icon="📧"
)

# Function to make the application speak the text
def speak(text):
    pythoncom.CoInitialize()  # Initialize COM library
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(text)

# Load the model and vectorizer
model = pickle.load(open('spam.pkl', 'rb'))
cv = pickle.load(open('vectorizer.pkl', 'rb'))

# Function to add background image
def add_bg_from_url():
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url("https://www.publicdomainpictures.net/pictures/320000/velka/background-image.png");
            background-attachment: fixed;
            background-size: cover;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

# Function to set custom font and styling
def set_custom_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600&display=swap');

        html, body, [class*="css"]  {
            font-family: 'Poppins', sans-serif;
        }
        .main {
            background-color: rgba(255, 255, 255, 0.8);
            padding: 2rem;
            border-radius: 8px;
        }
        .stButton > button {
            background-color: gray;
            color: white;
            border-radius: 8px;
            border: none;
            font-size: 18px;
            padding: 10px 24px;
        }
        .stButton > button:hover {
            background-color: #00796B;
            color: #e6e6e6;
        }
        .footer {
            font-size: 0.8rem;
            position: fixed;
            left: 0;
            bottom: 0;
            width: 100%;
            background-color: #333;
            color: white;
            text-align: center;
            padding: 1rem 0;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

# Initialize session state for storing history
if 'history' not in st.session_state:
    st.session_state['history'] = []

# Function to clear history
def clear_history():
    st.session_state['history'] = []

# Main function to build the app
def main():
    # Adding background image and custom styling
    add_bg_from_url()
    set_custom_css()

    # Application header with styling
    st.title("📧 Email Spam Detector Application")
    st.markdown("### Built with Streamlit & Python")
    st.markdown("Developed by **Anurag Pandey**")
    st.write("---")

    # Sidebar for navigation
    st.sidebar.title("Navigation")
    activities = ["Classification", "History", "About"]
    choice = st.sidebar.radio("Go to", activities)

    # Classification section
    if choice == "Classification":
        st.markdown("Enter the email content below to check if it's spam or not.")
        msg = st.text_area("Enter the email content here", height=200)

        if st.button("Classify"):
            if msg:
                data = [msg]
                vec = cv.transform(data).toarray()
                result = model.predict(vec)

                # Determine the result and update history
                if result[0] == 0:
                    classification = "Not Spam"
                    st.success("✔️ This is Not A Spam Email")
                    speak("This is Not A Spam Email")
                else:
                    classification = "Spam"
                    st.error("❌ This is A Spam Email")
                    speak("This is A Spam Email")

                # Save to session state history
                st.session_state['history'].append({'message': msg, 'classification': classification})

            else:
                st.warning("Please enter email content to classify.")

    # History section
    elif choice == "History":
        st.subheader("📜 Classification History")
        if st.session_state['history']:
            for i, record in enumerate(st.session_state['history'], 1):
                st.write(f"**{i}. Email:** {record['message']}")
                st.write(f"**Classification:** {record['classification']}")
                st.write("---")

            # Button to clear history
            if st.button("Clear History"):
                clear_history()
                st.success("History has been cleared.")

        else:
            st.info("No classification history available.")

    # About section
    elif choice == "About":
        st.subheader("About This Application")
        st.markdown("""
        This application uses a machine learning model to classify emails as spam or not spam. It leverages 
        the power of Natural Language Processing (NLP) to analyze the content of emails.
        
        ### How It Works
        - Enter the content of an email in the text area provided.
        - Click on 'Classify' to see if the email is spam or not.
        - The model used is trained on a large dataset of emails and is quite accurate.
        
        ### Developer
        Developed by **Anurag Pandey**. Connect with me on [LinkedIn](https://www.linkedin.com).
        """)

    # Footer section
    st.markdown(
        """
        <div class="footer">
        <p>Developed by Anurag Pandey | <a href="https://www.linkedin.com" target="_blank">LinkedIn</a></p>
        </div>
        """,
        unsafe_allow_html=True
    )

# Run the main function
if __name__ == "__main__":
    main()
