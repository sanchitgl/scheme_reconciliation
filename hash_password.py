import pickle 
from pathlib import Path 
import streamlit_authenticator as stauth

password = "amit"

hashed_passwords = stauth.Hasher([password]).generate()

print(hashed_passwords)