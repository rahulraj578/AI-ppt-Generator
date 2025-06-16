Summary
This project automatically generates PowerPoint presentations using OpenAI’s Assistant API and DALL·E. It creates slides with titles, short content, and relevant AI-generated images—then compiles them into a .pptx file using Python.

⚙️ Technologies
Python 3

OpenAI GPT (Function Calling & Assistant API)

DALL·E for image generation

python-pptx, dotenv, Pillow

🚀 Setup
Clone the repo

Create a .env file with your OpenAI API key

Install dependencies:

bash
Copy
Edit
pip install -r requirements.txt
Run the script:

bash
Copy
Edit
python assistant-api-gpt4.py
📂 Output
.pptx file saved in powerpoint-ppt/

Each slide includes:

Title

Brief content (≤20 words)

AI-generated image (DALL·E)
