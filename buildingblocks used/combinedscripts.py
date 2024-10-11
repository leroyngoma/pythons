import os
import requests
import json
import win32com.client  # For accessing Outlook
import re  # Regular expressions for text extraction

# ----------------------------------------------
# Step 1: Extract text from Outlook emails
# ----------------------------------------------
def get_email_text():
    # Initialize Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Access inbox folder (Folder 6 is typically the inbox)
    inbox = outlook.GetDefaultFolder(6)
    
    # Get all emails from the inbox
    messages = inbox.Items
    
    # Initialize a variable to store the extracted prompt
    extracted_prompt = None

    # Loop through all messages
    for message in messages:
        try:
            # Check if 'extractstart' is in the message body
            if 'extractstart' in message.Body.lower():
                # Extract the text between 'extractstart' and 'extractend'
                match = re.search(r'extractstart(.*?)extractend', message.Body, re.DOTALL)
                if match:
                    extracted_prompt = match.group(1).strip()  # Get the extracted text and remove any extra spaces
                    break  # Stop after finding the first relevant email
        except Exception as e:
            print(f"Error reading message: {e}")
    
    return extracted_prompt


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
            "content": "You are a helpful and insightful AI assistant that analyzes survey responses to generate comprehensive AI readiness reports. Your goal is to provide users with actionable insights into their understanding of AI and its potential application in their business..."
        },
        {"role": "user", "content": extracted_prompt}
    ]
    
    # Payload for the request
    payload = {
        "messages": messages,
        "temperature": 0.7,
        "top_p": 0.95,
        "max_tokens": 2400
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
# Step 3: Main function to tie everything together
# ----------------------------------------------
def main():
    # Extract the prompt from email
    prompt = get_email_text()
    
    if prompt:
        print(f"Extracted Prompt: {prompt}")
        
        # Send the extracted prompt to the LLM and get the response
        llm_response = generate_response(prompt)
        
        if llm_response:
            print(f"LLM Response: {llm_response}")
        else:
            print("Error: LLM did not return a response.")
    else:
        print("No valid email found with 'extractstart' and 'extractend'.")


# ----------------------------------------------
# Run the combined process
# ----------------------------------------------
if __name__ == "__main__":
    main()
