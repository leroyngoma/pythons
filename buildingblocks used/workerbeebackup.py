import os
import requests
import json
import win32com.client  # For accessing Outlook
import re  # Regular expressions for text extraction
from reportlab.lib.pagesizes import letter  # For generating PDF with reportlab
from reportlab.pdfgen import canvas
import datetime
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.fonts import tt2ps
import pythoncom

from reportlab.lib import fonts

# ----------------------------------------------
# Step 1: Extract text and sender's email from Outlook emails
# ----------------------------------------------
def get_email_text_and_sender():
    # Initialize Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Access inbox folder (Folder 6 is typically the inbox)
    inbox = outlook.GetDefaultFolder(6)
    
    # Get all emails from the inbox
    messages = inbox.Items
    
    # Initialize variables to store the extracted prompt and sender's email
    extracted_prompt = None
    sender_email = None

    # Loop through all messages
    for message in messages:
        try:
            # Check if 'extractstart' is in the message body
            if 'extractstart' in message.Body.lower():
                # Extract the text between 'extractstart' and 'extractend'
                match = re.search(r'extractstart(.*?)extractend', message.Body, re.DOTALL)
                if match:
                    extracted_prompt = match.group(1).strip()  # Get the extracted text and remove any extra spaces
                    sender_email = message.SenderEmailAddress  # Extract the sender's email address
                    break  # Stop after finding the first relevant email
        except Exception as e:
            print(f"Error reading message: {e}")
    
    return extracted_prompt, sender_email


# ----------------------------------------------
# Step 2: LLM Interaction via Azure OpenAI
# ----------------------------------------------
API_KEY = "a3acda2a38d948e6b324f20bd3154e0e7"  # Replace with your actual API key
ENDPOINT = "https://readinessreporter.openai.azure.com/openai/deployments/gpt-4o-mini/chat/completions?api-version=2024-02-15-preview"

headers = {
    "Content-Type": "application/json",
    "api-key": API_KEY,
}

# Function to send the extracted prompt to the Azure OpenAI LLM API
def generate_response(extracted_prompt):
    # Initialize messages with the system prompt and extracted prompt
    messages = [
        {
            "role": "system",
            "content": "You are a helpful and insightful AI assistant that analyzes survey responses to generate comprehensive AI readiness reports.  Your goal is to provide users with actionable insights into their understanding of AI and its potential application in their business. \n\nYou will receive a series of survey questions and a user's corresponding answers. Using this information, you will draft a Full Report.  \n\n\n*   This report is for **AiBizHive** to gain deeper insight into the user's needs. \n*   Provide specific details and observations based on the user's survey responses.\n*   Maintain a neutral and objective tone.\n*   Organize your analysis into the following sections:\n    *   Understanding of AI and Its Applications\n    *   Data Readiness\n    *   AI Expertise and Resources\n    *   AI Project Management and Implementation\n    *   AI Investment Approach\n*   Within each section, provide concrete examples of how AiBizHive can assist the user based on their responses to the survey questions. \n\n**Survey Questions:**\n\nquestion 1\nWhat are some of the biggest challenges your business faces daily?\n1. We don’t have a clear idea of the challenges AI could help with.\n2. We face challenges but haven’t considered AI as a solution.\n3. We can identify some challenges, but not sure if AI can solve them.\n4. We know some areas AI could improve but need more research.\n5. We have clear challenges and know AI can help solve them.\n\nquestion 2\nIf you could magically improve just one thing about how your business operates, what would it be?\n1. I’m not sure AI could help with the issues we face.\n2. I don’t think the area we want to improve relates to AI.\n3. The area we want to improve might involve AI but needs exploration.\n4. AI could help improve our top priorities, but we need more clarity.\n5. AI directly aligns with what we want to improve the most.\n\nquestion 3\nDoes your business regularly collect and store data?\n1. We don’t really collect or store data.\n2. We collect some data, but it’s not very well organized.\n3. We collect and store data, but it’s not always easily accessible.\n4. We collect and organize data, but there are areas to improve.\n5. We have strong, organized data collection and storage practices.\n\nquestion 4\nHow confident are you in the accuracy and completeness of your business data?\n1. We have low confidence in the accuracy and completeness of our data.\n2. We face some issues with our data quality but have no process for improvement.\n3. We have some confidence in our data quality but could improve in some areas.\n4. We are fairly confident in our data quality but not 100% sure.\n5. We are very confident that our data is accurate and complete.\n\nquestion 5\nHow familiar are you with AI and its potential applications?\n1. I’m not familiar with AI at all.\n2. I’ve heard of AI but don’t know much about it.\n3. I have a basic understanding of AI and its uses.\n4. I have a good understanding of AI and its business applications.\n5. I’m highly knowledgeable about AI and its practical uses in business.\n\nquestion 6\nDoes your business have in-house AI expertise or partnerships?\n1. We have no AI expertise or partnerships.\n2. We are exploring options but don’t have expertise yet.\n3. We have some AI expertise or partnerships, but it’s limited.\n4. We have adequate AI expertise or partnerships to start small projects.\n5. We have strong in-house AI expertise or partnerships.\n\nquestion 7\nIs your business willing to invest financially in AI projects?\n1. We don’t have any budget allocated for AI.\n2. We might consider it but don’t have a budget yet.\n3. We have a small budget allocated for AI projects.\n4. We have a moderate budget for AI projects.\n5. We have significant financial resources allocated to AI.\n\nquestion 8\nWhat is your business’s approach to AI project investment?\n1. We are very cautious and skeptical about AI investments.\n2. We are somewhat cautious but open to small AI investments.\n3. We are open to moderate AI investments with measurable outcomes.\n4. We are willing to invest in AI projects with potential benefits.\n5. We are very proactive and prioritize AI investments.\n\nquestion 9\nDo you have the necessary technical infrastructure to support AI initiatives?\n1. We don’t have the technical infrastructure for AI.\n2. We have limited technical infrastructure for AI.\n3. We have some infrastructure but need improvements.\n4. We have adequate infrastructure to support basic AI.\n5. We have advanced infrastructure in place for AI.\n\nquestion 10\nHow prepared is your business for AI implementation?\n1. We are not prepared at all for AI implementation.\n2. We are in the early stages of preparation.\n3. We have made some preparations but need more readiness.\n4. We are fairly prepared but need final adjustments.\n5. We are fully prepared and ready to implement AI.\n\nquestion 11\nIs your business aware of regulatory and ethical considerations surrounding AI?\n1. We are not aware of any AI regulations or ethical concerns.\n2. We have limited awareness of AI regulations or ethics.\n3. We are somewhat aware but need more guidance on AI regulations.\n4. We have good awareness but could use more clarity.\n5. We are fully aware and compliant with AI regulations and ethics.\n\nquestion 12\nDo you have a clear understanding of how AI might impact privacy and security in your business?\n1. We don’t understand how AI might affect privacy and security.\n2. We have limited understanding of AI’s impact on privacy/security.\n3. We are aware but need further insight into AI privacy/security concerns.\n4. We have a good understanding of AI’s privacy/security impacts.\n5. We are very confident in our understanding of AI’s privacy and security implications.\n\nquestion 13\nDoes your business have a clear plan for managing AI projects?\n1. We don’t have any AI project management plans.\n2. We are exploring options for AI project management.\n3. We have some plans but need clearer project management structures.\n4. We have a clear plan for small AI projects.\n5. We have a comprehensive plan for managing large AI projects.\n\n\nquestion 14\nHow confident are you in your team’s ability to execute AI projects successfully?\n1. We don’t have confidence in our team’s ability to execute AI projects.\n2. We have limited confidence in our team’s ability.\n3. We are somewhat confident but need more resources or training.\n4. We are fairly confident in our team’s ability to execute AI projects.\n5. We are very confident in our team’s ability to execute successful AI projects.\n\n\nONCE I PASTE THE USER RESPONSES AND DETAILS A REPORT MUST BE GENERATED"
        },
        {"role": "user", "content": extracted_prompt}
    ]
    
    # Payload for the request
    payload = {
        "messages": messages,
        "temperature": 0.7,
        "top_p": 0.95,
        "max_tokens": 8000
    }
    
    # Send request to the Azure OpenAI service
    try:
        response = requests.post(ENDPOINT, headers=headers, json=payload)
        response.raise_for_status()  # Raise an error for bad responses
        response_data = response.json()
        
        # Extract and return the assistant's reply
        bot_reply = response_data['choices'][0]['message']['content']
        return bot_reply
        
    except requests.RequestException as e:
        print(f"Failed to make the request. Error: {e}")
        return None


# ----------------------------------------------
# Step 3: Generate PDF from LLM response using reportlab
# ----------------------------------------------
def generate_pdf(llm_response):
    # Define the PDF filename
    pdf_filename = "AI_Readiness_Report.pdf"

    # Create a PDF document
    doc = SimpleDocTemplate(pdf_filename, pagesize=letter)
    
    # Create an empty list to hold the PDF content
    content = []
    
    # Define styles for the document
    styles = getSampleStyleSheet()
    
    # Title style
    title_style = ParagraphStyle(
        name="Title",
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=20,
        fontName="Helvetica-Bold"
    )
    
    # Normal text style
    normal_style = ParagraphStyle(
        name="Normal",
        fontSize=12,
        leading=15,
        alignment=TA_LEFT
    )
    
    # Bold text style
    bold_style = ParagraphStyle(
        name="Bold",
        fontSize=12,
        fontName="Helvetica-Bold",
        leading=15,
        alignment=TA_LEFT
    )
    
    # Add the title and timestamp
    title = Paragraph("AI Readiness Report", title_style)
    timestamp = Paragraph(f"Generated on: {datetime.datetime.now()}", normal_style)
    content.append(title)
    content.append(Spacer(1, 12))
    content.append(timestamp)
    content.append(Spacer(1, 12))
    
    # Split the LLM response into lines and process each line for bold formatting
    for line in llm_response.split('\n'):
        # Find text between ** and apply bold
        bold_parts = re.split(r'(\*\*.*?\*\*)', line)
        for part in bold_parts:
            if part.startswith('**') and part.endswith('**'):
                # Remove the asterisks and make the text bold
                part = part[2:-2]
                content.append(Paragraph(part, bold_style))
            else:
                # Regular text
                content.append(Paragraph(part, normal_style))
        content.append(Spacer(1, 12))  # Add space between paragraphs
    
    # Build the PDF with content
    doc.build(content)
    
    return pdf_filename

# ----------------------------------------------
# Step 4: Send the PDF via email
# ----------------------------------------------
def send_email_with_pdf(sender_email, pdf_filename):
    # Initialize Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application")
    
    # Create a new mail item
    mail = outlook.CreateItem(0)  # 0 means a Mail Item
    
    # Set email details
    mail.To = sender_email
    mail.Subject = "Your AI Readiness Report"
    mail.Body = "Please find attached the AI Readiness Report based on your input."
    
    # Attach the PDF file
    mail.Attachments.Add(os.path.abspath(pdf_filename))
    
    # Send the email
    mail.Send()
    print(f"Email sent to {sender_email} with attached PDF.")


# ----------------------------------------------
# Step 5: Main function to tie everything together
# ----------------------------------------------
def main():
    # Extract the prompt and sender email from the inbox
    prompt, sender_email = get_email_text_and_sender()
    
    if prompt and sender_email:
        print(f"Extracted Prompt: {prompt}")
        
        # Send the extracted prompt to the LLM and get the response
        llm_response = generate_response(prompt)
        
        if llm_response:
            print(f"LLM Response: {llm_response}")
            
            # Generate a PDF from the LLM response
            pdf_filename = generate_pdf(llm_response)
            
            # Send the PDF to the sender's email
            send_email_with_pdf(sender_email, pdf_filename)
        else:
            print("Error: LLM did not return a response.")
    else:
        print("No valid email found with 'extractstart' and 'extractend' or no sender email detected.")


# ----------------------------------------------
# Run the combined process
# ----------------------------------------------
if __name__ == "__main__":
    main()
