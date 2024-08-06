# PDF Converter and Merger 👷🏻‍♀️

Welcome to the **PDF Converter and Merger**! This Streamlit application allows you to effortlessly convert various document formats (Excel, Word, Images) into PDFs and merge them into a single PDF file. 

![Start_Merging!](app_thumbnail.png)

## 🌐 Try it Out

[Click here to access the live app](https://huggingface.co/spaces/2002SM2002/Converter_and_Merger#e2df659)

## 🚀 Features

- **Excel to PDF Conversion**: Convert your Excel files (.xlsx) to PDF while maintaining the original formatting.
- **Word to PDF Conversion**: Convert your Word documents (.docx) to PDF.
- **Image to PDF Conversion**: Convert image files (.png, .jpg, etc.) to PDF.
- **PDF Merging**: Merge multiple PDF files into one seamless document.

## 🛠️ Technologies Used

- **HuggingFace**: Deployment.
- **Streamlit**: Interactive web applications.
- **Pandas**: Data manipulation and analysis.
- **Matplotlib**: Plotting and converting images to PDFs.
- **ReportLab**: Creating and styling PDFs.
- **python-docx**: Reading DOCX files.
- **PyPDF2**: Merging PDFs.
- **Pillow**: Handling images.

## 📥 Installation

1. **Clone the repository:**

    ```bash
    git clone https://github.com/Zissi-Milstein/Converter-and-Merger.git
    cd Converter-and-Merger
    ```

2. **Create a virtual environment and activate it:**

    ```bash
    python3 -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3. **Install the dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

## 🎮 Usage

1. **Run the Streamlit application:**

    ```bash
    streamlit run app.py
    ```

2. **Upload your files:**
    - Click on "Choose files" and select your documents (Excel, Word, Images, or PDFs).
    - Choose the desired orientation (landscape or portrait) for Excel and Word files.

3. **Convert and Merge:**
    - The application will convert the selected files to PDFs and merge them.
    - Download the merged PDF using the "Download Merged PDF" button.


## 📝 Requirements

Ensure you have the following packages installed:

```plaintext
streamlit==1.17.0
pandas==2.0.2
matplotlib==3.7.1
reportlab==3.6.0
python-docx==0.8.11
PyPDF2==3.0.1
Pillow==10.0.0
openpyxl==3.1.2
```

## 🌟 Contributing

Contributions are welcome! Please fork the repository and submit a pull request.
