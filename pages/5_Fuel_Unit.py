import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta
import dropbox
import io

st.set_page_config(page_title="KTC | Fuel Unit", page_icon="description/logo.png")

st.title("Fuel Unit")