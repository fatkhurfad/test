#!/bin/bash

# Ganti dengan nama file Python kamu
APP_FILE="test4.py"

# Cek apakah streamlit sudah terinstall
if ! command -v streamlit &> /dev/null
then
    echo "❌ Streamlit belum terinstall. Jalankan: pip install streamlit"
    exit
fi

# Jalankan aplikasi
echo "🚀 Menjalankan aplikasi $APP_FILE ..."
streamlit run "$APP_FILE"
