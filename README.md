# 📄 Tender Document Analyzer & Chatbot

An AI-powered web application that analyzes tender documents in PDF format, extracts important information, and allows users to ask questions about the uploaded tender through a chatbot.

The application is designed to reduce the manual effort required to read lengthy tender documents and extract important details such as bid dates, EMD, scope of work, payment terms, warranty, delivery period, and other key requirements.

## 🚀 Features

- 📤 Upload tender documents in PDF format
- 📑 Extract important information from tender documents
- 🤖 Ask questions about the uploaded tender using a chatbot
- 🔍 Search and retrieve relevant information from the document
- 📊 Automatically populate extracted information into an Excel template
- 📝 Extract important tender fields such as:
  - Company
  - Tender Number
  - GeM Number
  - Bid Date
  - EMD
  - Scope of Work
  - Warranty
  - CPBG
  - Payment Terms
  - Delivery Period
  - Area of Implementation
  - Package in a Nutshell
- 💻 Simple and interactive Streamlit interface

## 🛠️ Technologies Used

- **Python**
- **Streamlit**
- **PDF Processing**
- **Natural Language Processing (NLP)**
- **LLM / Generative AI**
- **Excel / XLSX**
- **CSS**

## 📂 Project Structure

```text
Tender-Document/
│
├── app.py                  # Main Streamlit application
├── requirements.txt        # Python dependencies
├── README.md               # Project documentation
│
├── data/
│   └── sample_tender.pdf   # Sample tender document
│
├── templates/
│   └── tender_template.xlsx
│
└── output/
    └── extracted_tender.xlsx
```

> The exact folder structure may vary depending on the implementation.

## ⚙️ Installation

### 1. Clone the repository

```bash
git clone <your-repository-url>
cd Tender-Document
```

### 2. Create a virtual environment

```bash
python -m venv venv
```

Activate it on macOS/Linux:

```bash
source venv/bin/activate
```

On Windows:

```bash
venv\Scripts\activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

## ▶️ Run the Application

Start the Streamlit application using:

```bash
streamlit run app.py
```

The application will open in your browser.

## 🔄 How It Works

The application follows the following workflow:

```text
Tender PDF
    ↓
Upload Document
    ↓
PDF Text Extraction
    ↓
Document Processing
    ↓
Information Extraction
    ↓
Structured Tender Data
    ↓
Excel Template
```

At the same time, users can interact with the uploaded document through the chatbot:

```text
User Question
      ↓
Document Search / Retrieval
      ↓
Relevant Tender Content
      ↓
AI Processing
      ↓
Answer
```

## 📊 Extracted Information

The system is designed to identify important tender-related information including:

| Field | Description |
|---|---|
| Company | Organization issuing the tender |
| Tender No. | Tender identification number |
| GeM | Government e-Marketplace reference |
| Bid Date | Tender submission/bid deadline |
| EMD | Earnest Money Deposit |
| Scope | Scope of work |
| Warranty | Warranty requirements |
| CPBG | Contract Performance Bank Guarantee |
| Payment Terms | Payment conditions |
| Delivery Period | Required delivery/completion period |
| Area of Implementation | Location/implementation area |
| Package in a Nutshell | Brief summary of the tender package |

## 💬 Tender Chatbot

The chatbot allows users to ask questions directly about the uploaded tender document.

Example questions:

```text
What is the EMD amount?

What is the last date for bid submission?

What is the scope of work?

What is the warranty period?

What are the payment terms?

What is the delivery period?

Where will the project be implemented?
```

The system uses the uploaded tender document as the source of information for answering these questions.

## 📥 Excel Output

After processing the document, the extracted information can be organized into an Excel template.

Example:

| Company | Tender No | EMD | Bid Date | Warranty | Payment Terms |
|---|---|---|---|---|---|
| ABC Ltd. | TENDER-001 | ₹5,00,000 | 10/12/2026 | 2 Years | As per tender |

This helps convert unstructured tender documents into structured, usable data.

## 🎯 Use Cases

This application can be useful for:

- Tender management teams
- Procurement departments
- Government contractors
- Construction companies
- Engineering companies
- Business development teams
- Organizations handling large numbers of tenders

## 🔮 Future Improvements

Possible future enhancements include:

- Support for multiple PDF uploads
- OCR support for scanned PDFs
- Improved extraction accuracy
- Automatic tender comparison
- Tender deadline alerts
- Advanced document search
- Multiple Excel template support
- Automatic tender eligibility analysis
- Tender summarization
- Integration with cloud storage
- User authentication and document history

## ⚠️ Limitations

The accuracy of extracted information depends on:

- Quality and structure of the uploaded PDF
- Whether the document contains selectable text
- Formatting of the tender document
- Availability of the required information
- Compatibility between the uploaded tender and the selected Excel template

For scanned or poorly formatted documents, OCR and additional preprocessing may be required.

## 👩‍💻 Author

**Himali Sharma**

B.Tech – Computer Science and Engineering  
Vellore Institute of Technology (VIT)

---

## 📜 License

This project is intended for educational, research, and demonstration purposes.
