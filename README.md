# Wild Bot
The purpose of this scrypt was to simplify the booking process for Wild Entertainment. It takes a new entertainer request recieved via email, determines the availability of the entertainers requested, then sends an SMS to the entertainers using an exploit I found in Google Voice with Twilio integration. The exploit I found is the highlight of this project.

Workflow:
1. New Entertainer Request (N.E.R.) are recieved.
2. N.E.R. are quarantined to their own folder to be processed individually.
3. The most recent N.E.R. is scraped and has it's data stored in variables. 
4. The bot takes the names of the entertainers requested, the date and time of the event and compares it with the Microsoft Exchange calendar. 
5. If an entertainer is available, they are labeled as "OPEN" and the bot will send them an SMS with all the event information.
6. After the bot has contacted or labeled each entertainer requested, it creates an event in the Microsoft Exchange calendar at the specific date and time of the actual event. The subject line reflects who the customer is, who they have requested and who is available/been contacted.

The biggest accomplishment of this project is the Google Voice exploit I found. I have figured out how to use the Gooogle Voice userinterface with programmable SMS provided by Twilio. This essentially gives Google Voice an API.

To use Google Voice GUI for Twilio:
1. Have your GV number forward SMS to your twilio number and your personal phone number.
2. Have your contacts send an SMS to your Google voice number.
3. On your personal phone, you will recieve your contact's SMS from an unknown phone number. This phone number is assigned by Google Voice and in my testing does not change. This phone number represents the "google voice id" of your contacts.
4. To send an SMS through Google voice via Twilio API, simply send your sms via the Twilio API to the "google voice ID" from step 3.


