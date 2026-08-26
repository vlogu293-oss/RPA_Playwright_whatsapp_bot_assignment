# Bot workflow
Start Chrome with a persistent WhatsApp profile.
Open WhatsApp Web.
If not already logged in, pause and ask you to scan the QR code.
Read contacts.xlsx.
For every contact:
Validate the phone number.
Open/search the WhatsApp chat.
Replace {name} with the contact's name.
Send the message.
Wait for the sent message to appear.
Take a screenshot.
Extract the last 3 incoming messages from that contact.
Store results such as:
Name
Phone
Message sent
Sent status
Screenshot filename
Last 3 messages
Timestamp
Error, if any
Generate:
whatsapp_report.json
whatsapp_report.xlsx

For example, the JSON record can look like:

{
    "name": "Ravi",
    "phone": "+919876543210",
    "message_sent": "Hello Ravi, this is a test message.",
    "status": "sent",
    "timestamp": "2026-08-26 20:45:12",
    "screenshot": "screenshots/Ravi_919876543210.png",
    "last_3_messages": [
        "Okay, thank you",
        "I will check it",
        "Sure"
    ],
    "error": null
}
Important implementation detail

For WhatsApp Web, I recommend using a persistent Chrome user profile rather than trying to automate the QR-code login. That means you scan the QR code manually once, and subsequent runs can reuse the logged-in session.