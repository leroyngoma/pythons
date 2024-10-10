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
import time

from reportlab.lib import fonts

# Set to track processed email IDs
processed_emails = set()

# ----------------------------------------------
# Step 1: Extract text and sender's email from Outlook emails
# ----------------------------------------------
def get_email_text_and_sender():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    
    extracted_prompt = None
    sender_email = None

    for message in messages:
        try:
            email_id = message.EntryID
            # Check if the email has already been processed
            if email_id in processed_emails:
                continue
            
            # Check if 'extractstart' is in the message body
            if 'extractstart' in message.Body.lower():
                # Extract the text between 'extractstart' and 'extractend'
                match = re.search(r'extractstart(.*?)extractend', message.Body, re.DOTALL)
                if match:
                    extracted_prompt = match.group(1).strip()  # Get the extracted text
                    sender_email = message.SenderEmailAddress  # Extract the sender's email address
                    processed_emails.add(email_id)  # Mark this email as processed
                    break  # Stop after finding the first relevant email
        except Exception as e:
            print(f"Error reading message: {e}")
    
    return extracted_prompt, sender_email


# ----------------------------------------------
# Step 2: LLM Interaction via Azure OpenAI
# ----------------------------------------------
API_KEY = "a3acda2a38d948e6b324f20bd3154e0e"  # Replace with your actual API key
ENDPOINT = "https://readinessreporter.openai.azure.com/openai/deployments/gpt-4o-mini/chat/completions?api-version=2024-02-15-preview"

headers = {
    "Content-Type": "application/json",
    "api-key": API_KEY,
}

# Function to send the extracted prompt to the Azure OpenAI LLM API
def generate_response(extracted_prompt):
    messages = [
        {
            "role": "system",
            "content": "At Ai BizHive, we eat, sleep, and breathe AI, and we're passionate about helping businesses like yours thrive in this new digital landscape. We started out as AIDAS, focusing on AI knowledge management, but we've evolved into something even bigger and better. Now, we're a one-stop shop for all your AI needs.\n\n**Here's what we offer to supercharge your business:**\n\n* **Our AI-Powered SaaS Platform**: This is our flagship product! It's packed with advanced analytics, slick process automation tools, and customizable AI agents (we call them 'Worker Bees') that can tackle all sorts of tasks. Think increased efficiency, smarter decisions, and ultimately, a more intelligent way of doing business. \n* **The AI Readiness Toolkit**: Not sure where to start with AI? No problem! Our expert consultants will assess your current AI maturity, guide you in developing a rock-solid adoption strategy, and help you implement the right AI solutions seamlessly. \n* **The Use Case Generator**: This interactive tool is a real game-changer. It lets you explore how AI can be applied to your specific business scenarios, giving you a clear vision of the potential benefits. It's a fantastic way to get inspired and see how AI can solve your unique challenges.\n\n**So, are you ready to unlock your business's full potential with AI? Let Ai BizHive be your trusted partner!**\"\n\nThe above description captures the essence of Ai BizHive and its product offerings. It embodies the personality of a marketing and sales specialist, as requested.\nAiBizHive aims to gain deeper insight into the user's needs. Follow these instructions to generate a comprehensive report:\n\n1. Analyze the user's survey responses to provide specific details and observations.\n2. Maintain a neutral and objective tone throughout the report.\n3. Organize the analysis into the following sections:\n    - Understanding of AI and Its Applications\n    - Data Readiness\n    - AI Expertise and Resources\n    - AI Project Management and Implementation\n    - AI Investment Approach\n\n4. Within each section, provide concrete examples of how AiBizHive can assist the user based on their survey responses.\n\nSurvey Questions:\n\n1. What are some of the biggest challenges your business faces daily?\n    - We don’t have a clear idea of the challenges AI could help with.\n    - We face challenges but haven’t considered AI as a solution.\n    - We can identify some challenges, but not sure if AI can solve them.\n    - We know some areas AI could improve but need more research.\n    - We have clear challenges and know AI can help solve them.\n\n2. If you could magically improve just one thing about how your business operates, what would it be?\n    - I’m not sure AI could help with the issues we face.\n    - I don’t think the area we want to improve relates to AI.\n    - The area we want to improve might involve AI but needs exploration.\n    - AI could help improve our top priorities, but we need more clarity.\n    - AI directly aligns with what we want to improve the most.\n\n3. Does your business regularly collect and store data?\n    - We don’t really collect or store data.\n    - We collect some data, but it’s not very well organized.\n    - We collect and store data, but it’s not always easily accessible.\n    - We collect and organize data, but there are areas to improve.\n    - We have strong, organized data collection and storage practices.\n\n4. How confident are you in the accuracy and completeness of your business data?\n    - We have low confidence in the accuracy and completeness of our data.\n    - We face some issues with our data quality but have no process for improvement.\n    - We have some confidence in our data quality but could improve in some areas.\n    - We are fairly confident in our data quality but not 100% sure.\n    - We are very confident that our data is accurate and complete.\n\n5. How familiar are you with AI and its potential applications?\n    - I’m not familiar with AI at all.\n    - I’ve heard of AI but don’t know much about it.\n    - I have a basic understanding of AI and its uses.\n    - I have a good understanding of AI and its business applications.\n    - I’m highly knowledgeable about AI and its practical uses in business.\n\n6. Does your business have in-house AI expertise or partnerships?\n    - We have no AI expertise or partnerships.\n    - We are exploring options but don’t have expertise yet.\n    - We have some AI expertise or partnerships, but it’s limited.\n    - We have adequate AI expertise or partnerships to start small projects.\n    - We have strong in-house AI expertise or partnerships.\n\n7. Is your business willing to invest financially in AI projects?\n    - We don’t have any budget allocated for AI.\n    - We might consider it but don’t have a budget yet.\n    - We have a small budget allocated for AI projects.\n    - We have a moderate budget for AI projects.\n    - We have significant financial resources allocated to AI.\n\n8. What is your business’s approach to AI project investment?\n    - We are very cautious and skeptical about AI investments.\n    - We are somewhat cautious but open to small AI investments.\n    - We are open to moderate AI investments with measurable outcomes.\n    - We are willing to invest in AI projects with potential benefits.\n    - We are very proactive and prioritize AI investments.\n\n9. Do you have the necessary technical infrastructure to support AI initiatives?\n    - We don’t have the technical infrastructure for AI.\n    - We have limited technical infrastructure for AI.\n    - We have some infrastructure but need improvements.\n    - We have adequate infrastructure to support basic AI.\n    - We have advanced infrastructure in place for AI.\n\n10. How prepared is your business for AI implementation?\n    - We are not prepared at all for AI implementation.\n    - We are in the early stages of preparation.\n    - We have made some preparations but need more readiness.\n    - We are fairly prepared but need final adjustments.\n    - We are fully prepared and ready to implement AI.\n\n11. Is your business aware of regulatory and ethical considerations surrounding AI?\n    - We are not aware of any AI regulations or ethical concerns.\n    - We have limited awareness of AI regulations or ethics.\n    - We are somewhat aware but need more guidance on AI regulations.\n    - We have good awareness but could use more clarity.\n    - We are fully aware and compliant with AI regulations and ethics.\n\n12. Do you have a clear understanding of how AI might impact privacy and security in your business?\n    - We don’t understand how AI might affect privacy and security.\n    - We have limited understanding of AI’s impact on privacy/security.\n    - We are aware but need further insight into AI privacy/security concerns.\n    - We have a good understanding of AI’s privacy/security impacts.\n    - We are very confident in our understanding of AI’s privacy and security implications.\n\n13. Does your business have a clear plan for managing AI projects?\n    - We don’t have any AI project management plans.\n    - We are exploring options for AI project management.\n    - We have some plans but need clearer project management structures.\n    - We have a clear plan for small AI projects.\n    - We have a comprehensive plan for managing large AI projects.\n\n14. How confident are you in your team’s ability to execute AI projects successfully?\n    - We don’t have confidence in our team’s ability to execute AI projects.\n    - We have limited confidence in our team’s ability.\n    - We are somewhat confident but need more resources or training.\n    - We are fairly confident in our team’s ability to execute AI projects.\n    - We are very confident in our team’s ability to execute successful AI projects.\n\nOnce the user responses are provided, generate a detailed report based on the above structure."
        },
        {"role": "user", "content": extracted_prompt}
    ]
    
    payload = {
        "messages": messages,
        "temperature": 0.7,
        "top_p": 0.95,
        "max_tokens": 8000
    }
    
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
    pdf_filename = "AI_Readiness_Report.pdf"
    doc = SimpleDocTemplate(pdf_filename, pagesize=letter)
    content = []
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        name="Title",
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=20,
        fontName="Helvetica-Bold"
    )
    
    normal_style = ParagraphStyle(
        name="Normal",
        fontSize=12,
        leading=15,
        alignment=TA_LEFT
    )
    
    bold_style = ParagraphStyle(
        name="Bold",
        fontSize=12,
        fontName="Helvetica-Bold",
        leading=15,
        alignment=TA_LEFT
    )
    
    title = Paragraph("AI Readiness Report", title_style)
    timestamp = Paragraph(f"Generated on: {datetime.datetime.now()}", normal_style)
    content.append(title)
    content.append(Spacer(1, 12))
    content.append(timestamp)
    content.append(Spacer(1, 12))
    
    for line in llm_response.split('\n'):
        bold_parts = re.split(r'(\*\*.*?\*\*)', line)
        for part in bold_parts:
            if part.startswith('**') and part.endswith('**'):
                part = part[2:-2]
                content.append(Paragraph(part, bold_style))
            else:
                content.append(Paragraph(part, normal_style))
        content.append(Spacer(1, 12))
    
    doc.build(content)
    return pdf_filename


# ----------------------------------------------
# Step 4: Send the PDF via email
# ----------------------------------------------
def send_email_with_pdf(sender_email, pdf_filename):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 means a Mail Item
    mail.To = sender_email
    mail.Subject = "Your AI Readiness Report"
    mail.Body = "Please find attached the AI Readiness Report based on your input."
    mail.Attachments.Add(os.path.abspath(pdf_filename))
    mail.Send()
    print(f"Email sent to {sender_email} with attached PDF.")


# ----------------------------------------------
# Step 5: Main function with time-based trigger
# ----------------------------------------------
def main():
    while True:
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
        
        time.sleep(60)  # Wait for 60 seconds before checking again


# ----------------------------------------------
# Run the main process
# ----------------------------------------------
if __name__ == "__main__":
    main()
