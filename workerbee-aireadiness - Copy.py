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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.units import inch, cm
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
    extracted_email = None

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
                    
                    # Search for an email address within the extracted text
                    email_match = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', extracted_prompt)
                    if email_match:
                        extracted_email = email_match.group(0)  # Extract the found email address
                    
                    processed_emails.add(email_id)  # Mark this email as processed
                    break  # Stop after finding the first relevant email
        except Exception as e:
            print(f"Error reading message: {e}")
    
    return extracted_prompt, extracted_email


# ----------------------------------------------
# Step 2: LLM Interaction via Azure OpenAI
# ----------------------------------------------
API_KEY = "a3acda2a38d948"  # Replace with your actual API key
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
      "content": [
        {
          "type": "text",
          "text": "Always make each paragraph atleast 4 sentences long. Ai BizHive Your AI Business Partner Hi there! Im Ai BizHive, your\ndedicated AI guide. Im all about boosting your business with the power of AI – think\nof me as your marketing and sales specialist, but with a digital brain! Ai BizHive,\nwe eat, sleep, and breathe AI, and we're passionate about helping businesses like\nyours thrive in this new digital landscape. We started out as AIDAS, focusing on AI\nknowledge management, but weve evolved into something even bigger and better.\nNow, we&#39;re a one-stop shop for all your AI needs. Heres what we offer to\nsupercharge your business Our AI-Powered SaaS Platform: This is our flagship\nproduct! Its packed with advanced analytics, slick process automation tools, and\ncustomizable AI agents (we call them Worker Bees) that can tackle all sorts of\ntasks. Think increased efficiency, smarter decisions, and ultimately, a more intelligent\n\nway of doing business. The AI Readiness Toolkit: Not sure where to start with AI?\nNo problem! Our expert consultants will assess your current AI maturity, guide you in\ndeveloping a rock-solid adoption strategy, and help you implement the right AI\nsolutions seamlessly. The Use Case Generator: This interactive tool is a real\ngame-changer. It lets you explore how AI can be applied to your specific business\nscenarios, giving you a clear vision of the potential benefits. Its a fantastic way to get\ninspired and see how AI can solve your unique challenges. So, are you ready to\nunlock your business's full potential with AI? Let Ai BizHive be your trusted\npartner!"
        }
      ]
    },
    {
      "role": "user",
      "content": [
        {
          "type": "text",
          "text": "You got a new message from Ai BizHive: User name Leroy\nNgoma email address: leroyngoma@gmail.com\\norganisation name:\naibizhive\\norganisation size: 1-10User score: question 1\\nWhat are some of\nthe biggest challenges your business faces daily?\\n1. We don’t have a clear idea of\nthe challenges AI could help with.\\n2. We face challenges but haven’t considered AI\nas a solution.\\n3. We can identify some challenges, but not sure if AI can solve\nthem.\\n4. We know some areas AI could improve but need more research.\\n5. We\nhave clear challenges and know AI can help solve them.\\nUser selected:\n5\\n\\nquestion 2\\nIf you could magically improve just one thing about how your\nbusiness operates, what would it be?\\n1. I’m not sure AI could help with the issues\nwe face.\\n2. I don’t think the area we want to improve relates to AI.\\n3. The area we\nwant to improve might involve AI but needs exploration.\\n4. AI could help improve\nour top priorities, but we need more clarity.\\n5. AI directly aligns with what we want to\nimprove the most.\\nUser selected: 1\\n\\nquestion 3\\nDoes your business regularly\ncollect and store data?\\n1. We don’t really collect or store data.\\n2. We collect some\ndata, but it’s not very well organized.\\n3. We collect and store data, but it’s not\nalways easily accessible.\\n4. We collect and organize data, but there are areas to\nimprove.\\n5. We have strong, organized data collection and storage practices.\\nUser\nselected: 5\\n\\nquestion 4\\nHow confident are you in the accuracy and completeness\nof your business data?\\n1. We have low confidence in the accuracy and\ncompleteness of our data.\\n2. We face some issues with our data quality but have\nno process for improvement.\\n3. We have some confidence in our data quality but\n\ncould improve in some areas.\\n4. We are fairly confident in our data quality but not\n100% sure.\\n5. We are very confident that our data is accurate and complete.\\nUser\nselected: 1\\n\\nquestion 5\\nHow familiar are you with AI and its potential\napplications?\\n1. I’m not familiar with AI at all.\\n2. I’ve heard of AI but don’t know\nmuch about it.\\n3. I have a basic understanding of AI and its uses.\\n4. I have a good\nunderstanding of AI and its business applications.\\n5. I’m highly knowledgeable\nabout AI and its practical uses in business.\\nUser selected: 5\\n\\nquestion 6\\nDoes\nyour business have in-house AI expertise or partnerships?\\n1. We have no AI\nexpertise or partnerships.\\n2. We are exploring options but don’t have expertise\nyet.\\n3. We have some AI expertise or partnerships, but it’s limited.\\n4. We have\nadequate AI expertise or partnerships to start small projects.\\n5. We have strong in-\nhouse AI expertise or partnerships.\\nUser selected: 1\\n\\nquestion 7\\nIs your\nbusiness willing to invest financially in AI projects?\\n1. We don’t have any budget\nallocated for AI.\\n2. We might consider it but don’t have a budget yet.\\n3. We have a\nsmall budget allocated for AI projects.\\n4. We have a moderate budget for AI\nprojects.\\n5. We have significant financial resources allocated to AI.\\nUser selected:\n5\\n\\nquestion 8\\nWhat is your business’s approach to AI project investment?\\n1. We\nare very cautious and skeptical about AI investments.\\n2. We are somewhat\ncautious but open to small AI investments.\\n3. We are open to moderate AI\ninvestments with measurable outcomes.\\n4. We are willing to invest in AI projects\nwith potential benefits.\\n5. We are very proactive and prioritize AI\ninvestments.\\nUser selected: 1\\n\\nquestion 9\\nDo you have the necessary technical\ninfrastructure to support AI initiatives?\\n1. We don’t have the technical infrastructure\nfor AI.\\n2. We have limited technical infrastructure for AI.\\n3. We have some\ninfrastructure but need improvements.\\n4. We have adequate infrastructure to\nsupport basic AI.\\n5. We have advanced infrastructure in place for AI.\\nUser\nselected: 5\\n\\nquestion 10\\nHow prepared is your business for AI\nimplementation?\\n1. We are not prepared at all for AI implementation.\\n2. We are in\nthe early stages of preparation.\\n3. We have made some preparations but need\nmore readiness.\\n4. We are fairly prepared but need final adjustments.\\n5. We are\nfully prepared and ready to implement AI.\\nUser selected: 1\\n\\nquestion 11\\nIs your\nbusiness aware of regulatory and ethical considerations surrounding AI?\\n1. We are\nnot aware of any AI regulations or ethical concerns.\\n2. We have limited awareness\nof AI regulations or ethics.\\n3. We are somewhat aware but need more guidance on\nAI regulations.\\n4. We have good awareness but could use more clarity.\\n5. We are\nfully aware and compliant with AI regulations and ethics.\\nUser selected:\n5\\n\\nquestion 12\\nDo you have a clear understanding of how AI might impact privacy\nand security in your business?\\n1. We don’t understand how AI might affect privacy\nand security.\\n2. We have limited understanding of AI’s impact on\nprivacy/security.\\n3. We are aware but need further insight into AI privacy/security\nconcerns.\\n4. We have a good understanding of AI’s privacy/security impacts.\\n5.\nWe are very confident in our understanding of AI’s privacy and security\nimplications.\\nUser selected: 1\\n\\nquestion 13\\nDoes your business have a clear\n\nplan for managing AI projects?\\n1. We don’t have any AI project management\nplans.\\n2. We are exploring options for AI project management.\\n3. We have some\nplans but need clearer project management structures.\\n4. We have a clear plan for\nsmall AI projects.\\n5. We have a comprehensive plan for managing large AI\nprojects.\\nUser selected: 5\\n\\nquestion 14\\nHow confident are you in your team’s\nability to execute AI projects successfully?\\n1. We don’t have confidence in our\nteam’s ability to execute AI projects.\\n2. We have limited confidence in our team’s\nability.\\n3. We are somewhat confident but need more resources or training.\\n4. We\nare fairly confident in our team’s ability to execute AI projects.\\n5. We are very\nconfident in our team’s ability to execute successful AI projects.\\nUser selected:\n1\\nBest wishes,\\nAi BizHive\\n\\"
        }
      ]
    },
    {
      "role": "assistant",
      "content": [
        {
          "type": "text",
          "text": "AI Readiness Report: Leroy Ngoma\\nemail address:\nleroyngoma@gmail.com\\norganisation name: aibizhive\\norganisation size: 1-\n10\\n\\nTotal Score: 41 out of 70\\nPercentage Readiness: 58.6%\\n\\nExecutive\nSummary\\n \\nLeroy, your responses to the AI Readiness Survey indicate a strong\nfoundational understanding of AI and its potential benefits for your business. You\nclearly recognize the transformative power of AI and are eager to explore its\napplications. Below are the key takeaways and areas for improvement.\\n\\nKey\nTakeaways\\n \\n\\nStrengths\\tOpportunities for Improvement\\nAdvanced AI\nKnowledge: High familiarity with AI.\\tBridging Vision and Execution: Identify specific\napplications of AI.\\nData-Driven Mindset: Strong data collection practices.\\tEnsuring\nData Accuracy: Validate and refine data practices.\\nCommitment to Investment:\nWilling to invest in AI.\\tNavigating Implementation: Support for expertise and project\nmanagement.\\n \\t\\n \\n\\nDetailed Report\\n \\n\\n1. Understanding of AI and Its\nApplications\\n \\n\\nStrength\\tOpportunity\\tHow AiBizHive Can Help\\nHigh familiarity\nwith AI and its potential uses.\\tExplore specific AI applications for unique business\nchallenges.\\t- Conduct needs assessments.\\n- Facilitate use case workshops.\\n-\nDevelop a customized AI roadmap.\\n \\t\\t\\n \\n\\n2. Data Readiness\\n\n\\n\\nStrength\\tArea for Improvement\\tHow AiBizHive Can Help\\nStrong data\ncollection and storage practices.\\tLow confidence in data accuracy and\ncompleteness.\\t- Conduct data quality assessments.\\n- Implement data cleansing\n\ntechniques.\\n- Establish data governance protocols.\\n \\t\\t\\n \\n\\n3. AI Expertise and\nResources\\n \\n\\nArea for Improvement\\tHow AiBizHive Can Help\\nLack of in-house\nAI expertise and partnerships.\\t- Talent acquisition support.\\n- Facilitate strategic\npartnerships.\\n- Provide hands-on training and workshops.\\n \\t\\n \\n\\n4. AI Project\nManagement and Implementation\\n \\n\\nStrength\\tArea for Improvement\\tHow\nAiBizHive Can Help\\nComprehensive project management plan in place.\\tNeed\nexperience to execute complex AI projects.\\t- Share project management best\npractices.\\n- Provide implementation support.\\n- Offer change management\nguidance.\\n \\t\\t\\n \\n\\n5. AI Investment Approach\\n \\n\\nStrength\\tArea for\nImprovement\\tHow AiBizHive Can Help\\nStrong willingness to invest in AI\nprojects.\\tCautious and skeptical approach may limit potential.\\t- Develop a robust\nbusiness case.\\n- Identify and prioritize high-value AI projects.\\n- Provide ongoing\nperformance measurement.\\n \\t\\t\\n \\n\\nNext Steps\\n \\nTo maximize the potential of\nAI for your business, consider the following next steps:\\n\\nSchedule a Consultation:\nDiscuss specific challenges and opportunities with an AiBizHive\nconsultant.\\nParticipate in Workshops: Engage in use case workshops to identify\nactionable AI applications.\\nInitiate Data Quality Assessment: Begin the process of\nensuring data accuracy and reliability.\\n \\n\\nConclusion\\n \\nLeroy, your readiness to\nexplore AI is commendable. With the right support from AiBizHive, you can\neffectively bridge the gap between AI knowledge and practical implementation,\npaving the way for transformative growth in your business.\\n\\nBest wishes,\\nAi\nBizHive\\n"
        }
      ]
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
        alignment=TA_LEFT,
        spaceAfter=10,
        fontName="Helvetica-Bold"
    )
    
    normal_style = ParagraphStyle(
        name="Normal",
        fontSize=12,
        leading=15,
        alignment=TA_LEFT
    )

    small_style = ParagraphStyle(
        name="small",
        fontSize=8,
        leading=5,
        alignment=TA_LEFT
    )
    
    bold_style = ParagraphStyle(
        name="Bold",
        fontSize=12,
        fontName="Helvetica-Bold",
        leading=15,
        alignment=TA_LEFT
    )
    logo_path = "logo.png"  # Path to your logo file
    try:
        logo = Image(logo_path, 4 * cm, 1 * cm)  
        logo.hAlign = 'LEFT'  
        content.append(logo)
        content.append(Spacer(1, 12))  # Add some space after the logo
    except Exception as e:
        print(f"Error loading logo: {e}")
    title = Paragraph("AI Readiness Report", title_style)
    timestamp = Paragraph(f"Generated on: {datetime.datetime.now()}", small_style)
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
            elif part.startswith('###'):
                part = part[3:]
                content.append(Paragraph(part, title_style))
            elif part.startswith('##'):
                part = part[2:]
                content.append(Paragraph(part, title_style))
            elif part.startswith('---'):
                part = part[3:]
                content.append(Paragraph(part, title_style))
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
    mail.Body = "Please find attached the AI Readiness Report based on your input. We would love to get in touch with you to discuss further"
    mail.Attachments.Add(os.path.abspath(pdf_filename))
    mail.Send()
    print(f"Email sent to {sender_email} with attached PDF.")


# ----------------------------------------------
# Step 5: Main function with time-based trigger
# ----------------------------------------------
def main():
    while True:
        prompt, extracted_email = get_email_text_and_sender()
        
        if prompt and extracted_email:
            print(f"Extracted Prompt: {prompt}")
            print(f"Extracted Email: {extracted_email}")
            
            # Send the extracted prompt to the LLM and get the response
            llm_response = generate_response(prompt)
            
            if llm_response:
                print(f"LLM Response: {llm_response}")
                
                # Generate a PDF from the LLM response
                pdf_filename = generate_pdf(llm_response)
                
                # Send the PDF to the extracted email address
                send_email_with_pdf(extracted_email, pdf_filename)
            else:
                print("Error: LLM did not return a response.")
        else:
            print("No valid email or prompt found with 'extractstart' and 'extractend'.")
        
        time.sleep(60)  # Wait for 60 seconds before checking again

# ----------------------------------------------
# Run the main process
# ----------------------------------------------
if __name__ == "__main__":
    main()
