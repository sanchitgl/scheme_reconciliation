import pickle 
from pathlib import Path 
import streamlit_authenticator as stauth

password = "admin"

hashed_passwords = stauth.Hasher([password]).generate()

print(hashed_passwords)