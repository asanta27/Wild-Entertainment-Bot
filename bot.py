from exchangelib import DELEGATE, Account, ServiceAccount, EWSDateTime, EWSTimeZone, CalendarItem, Configuration
import datetime
import time
from twilio.rest import Client
import html

# Twilio Credentials
account_sid = "#########"
auth_token = "#########""

client = Client(account_sid, auth_token)

# Exchange Credentials
credentials = ServiceAccount(username='#########"', password='#########"')
account = Account(primary_smtp_address='#########"', credentials=credentials, autodiscover=True, access_type=DELEGATE)

# Cache Exchange Credentials. This prevents server timeout
ews_url = account.protocol.service_endpoint
ews_auth_type = account.protocol.auth_type
primary_smtp_address = account.primary_smtp_address
config = Configuration(service_endpoint=ews_url, credentials=credentials, auth_type=ews_auth_type)

# Create account
myaccount = Account(primary_smtp_address=primary_smtp_address, config=config, autodiscover=False, access_type=DELEGATE)

# Exchange constants
tz = EWSTimeZone.timezone('America/Los_Angeles')
reservations_calendar = myaccount.root.get_folder_by_name('WE Reservations')
availability_calendar = myaccount.root.get_folder_by_name('Talent Availability')
main_inbox = myaccount.inbox.all()
template_inbox = myaccount.root.get_folder_by_name('Template')
template_scan = main_inbox.filter(subject='New Request!')
executed_inbox = myaccount.root.get_folder_by_name('Executed')


# These are "Google Voice ID" phone number's. Not actual phone numbers of entertainers. 
contact = {'ADMIN': "+15559075019",
           'Justin': "+17863534720",
           'Fabian': "+13137181197",
           'Steve': "+17862501412",
           'Chris': "+19726769249",
           'Savage': "+16175280744",
           'Santos': "+17862453319",
           'Ricky': "+16175280709",
           'Aston': "+16178710541",
           'Javier': "+12027093181",
           'Silk': "+19726768619",
           'Kaleb': "+12027093464",
           'Adam': "+15043023037",
           'Cisco': "+16178710824",
           'Zac': "+13137181809",
           'Antonio': "+16782087060",
           'Keanu': "+17863533333",
           'Chase': "+19726769897",
           'Simon': "+16178719439",
           'Sebastian': "+12675075059",
           'Roddy': "+17862501575",
           'Tiger': "+19726769330"}

# Search the Email Body
def search_email(s, first, last):
    try:
        start = s.index(first) + len(first)
        end = s.index(last, start)
        return s[start:end]
    except ValueError:
        return ""


# Search contact date base
def search_contacts(values, searchName):
    if searchName in values:
                return searchName
    else:
        return "False"

# Text Talent
def text_talent(list, position):
    client.messages.create(to=contact[list[position]],
                           from_="+15552996262",
                           body= type_entertainer + " request for you!" + "\n\n" +
                           sms_date + "\n" +
                            event_start_time + "\n" +
                            event_length + "\n" +
                           sms_location + "\n" +
                           "Women: " + event_numgirls + "\n" +
                                 "Men: " + event_numguys + monty + "\n\n" +
                                 "Are you available to do this?" + "\n-Bot")
    return

# Create Outlook Event
def create_event(customer, body, date, open_men, messaged,location):
    item = CalendarItem(account=myaccount,
                        folder=reservations_calendar,  # account & folder required.
                        subject="(N) " + customer + " • OPEN: " + open_men + messaged,
                        body=body,
                        start=date, # start & end required
                        end=date,
                        location=location)
    item.save()  # This gives the item an item_id and a changekey
  #  item.subject = "(N) " + customer
    item.save()  # When the items has an item_id, this will update the item

# Begin bot
loop = 0
while loop == 0:
    time.sleep(2)

    # Quarantine New Request
    template_scan = main_inbox.filter(subject='New Request!')
    for items in template_scan:
        items.move(template_inbox)
    to_do_list = template_inbox.all().count()

    # Process most recent request
    if to_do_list != 0:
        latest_request = template_inbox.filter(subject='New Request!').order_by('-datetime_received')[:1]
        for items in latest_request:
            s = html.unescape(items.body)
            htmlBody = items.body
            items.move(executed_inbox)

        # Scrape email and store data in variables
        end_string = "<"
        customer_name = search_email(s, "Name: ", end_string).strip()
        customer_phone = search_email(s, "Phone: ", end_string)
        customer_email = search_email(s, "Email: ", end_string)

        event_date = search_email(s, "Date of Party: ", end_string) + " " + search_email(s, "Start Time: ", end_string)
        event_start_time = search_email(s, "Start Time: ", end_string)
        event_location = search_email(s, "Location: ", end_string).strip()
        event_length = html.unescape(search_email(s, "Duration: ", end_string))
        event_numgirls = search_email(s, "# of Girls: ", end_string)
        event_numguys = search_email(s, "# of Guys: ", end_string)

        type_entertainer = search_email(s, "Type: ", end_string)
        num_entertainers = search_email(s, "# Entertainers: ", end_string)
        first_choice = search_email(s, "Choice #1: ", end_string).strip()
        second_choice = search_email(s, "Choice #2: ", end_string).strip()
        third_choice = search_email(s, "Choice #3: ", end_string).strip()
        fourth_choice = search_email(s, "Choice #4: ", end_string).strip()
        monty = search_email(s, "Add-Ons: ", end_string)
        special_request = search_email(s, "Special Request:", end_string)


        print("|--- Start: " + customer_name + " ---|")

        # Removes integers from location for SMS
        sms_location = ''.join(i for i in event_location if not i.isdigit()).lstrip()

        # adds line break to monty if it exist
        if not monty:
            monty = ""
        else:
            monty = "\n" + monty

        # Converts "1 Hour" into "1" "hour"
        timeVal,stringTime = event_length.split(" ")
        timeHours = 0
        timeMins = 0

        # assigns values to timeHours & timeMinutes variables. This is needed for EWS date format
        if "minutes" in stringTime:
            timeMins = int(timeVal)
            timeHours = 0
        elif "hour" in stringTime:
            timeHours = int(timeVal)
            timeMins = 0

        # date is start date&time, enddate is end date&time + 30 minute cushion
        date = datetime.datetime.strptime(event_date, "%d %b %Y %I:%M %p")
        sms_date = date.strftime("%A, %b %d")
        cushion = datetime.timedelta(minutes=29)
        duration = datetime.timedelta(hours=timeHours,minutes=timeMins)
        enddate = date + duration
        partyDate = date - cushion
        previousWeek = date - datetime.timedelta(days=90)
        nextWeek = date + datetime.timedelta(days=90)

        top_choices = [first_choice,second_choice,third_choice,fourth_choice]

        # fourth choice is optional on form
        if not fourth_choice:
           top_choices.remove(fourth_choice)

        talent = 0
        new_first_choice = True
        open_men = []
        messaged = ""

        # Check availability of entertainer, send SMS if they are available.
        for i, choice in enumerate(top_choices):

        		# search for an existing event with this entertainer
            search_event = reservations_calendar.filter(subject__contains=choice,start__range=(
            tz.localize(EWSDateTime(previousWeek.year,previousWeek.month,previousWeek.day)),
            tz.localize(EWSDateTime(enddate.year, enddate.month, enddate.day, enddate.hour, enddate.minute))),end__range=(
            tz.localize(EWSDateTime(partyDate.year, partyDate.month, partyDate.day, partyDate.hour , partyDate.minute)),
            tz.localize(EWSDateTime(nextWeek.year,nextWeek.month,nextWeek.day))))

            # search availability calendar to make sure entertainer is not in town or at their day job
            search_availability = availability_calendar.filter(subject__contains=choice,start__range=(
            tz.localize(EWSDateTime(previousWeek.year,previousWeek.month,previousWeek.day)),
            tz.localize(EWSDateTime(enddate.year, enddate.month, enddate.day, enddate.hour, enddate.minute))),end__range=(
            tz.localize(EWSDateTime(partyDate.year, partyDate.month, partyDate.day, partyDate.hour , partyDate.minute)),
            tz.localize(EWSDateTime(nextWeek.year,nextWeek.month,nextWeek.day))))
            
            final_date = tz.localize(EWSDateTime(date.year, date.month, date.day, date.hour, date.minute))
            event_subject = ""
            availability_subject = ""
            for item in search_event:
                event_subject = item.subject
            for item in search_availability:
                availability_subject = item.subject


            # Determines if entertainer is booked (C) or simply not available by the Subject lines of events found.
            if ("(C)" in event_subject) or (choice in availability_subject):
                top_choices[i] = choice + ": Booked"
                if "Party" in event_subject:
                    top_choices[i] = choice + ": Party Bus"
                print(top_choices[i])
            else:
                top_choices[i] = choice + ": Open"
                print(top_choices[i])

            if "Open" in top_choices[i]:
                open_men.insert(talent, top_choices[i][:-6])
                if search_contacts(contact, open_men[talent]) in top_choices[i]:
                    if new_first_choice is True:
                        print("> Sent SMS to " + open_men[talent] + " at " + time.strftime("%I:%M:%S %p"))
                        text_talent(open_men, talent)
                        messaged = " • TXT: " + open_men[talent]
                        new_first_choice = False
                elif "False" in search_contacts(contact, open_men[talent]):
                    if new_first_choice is True:
                        client.messages.create(to="+15559075019",
                                               from_="+15552996262",
                                               body="Please send the following message to " +
                                                     open_men[talent] + " for " + customer_name +
                                               "'s event.")

                        print(open_men[talent] + " is not in the database.")
                        print("> Sent Admin the SMS template at " + time.strftime("%I:%M:%S %p"))
                        admin = ['ADMIN']
                        time.sleep(1)
                        text_talent(admin, 0)
                        new_first_choice = False
                talent += 1

        create_event(customer_name, htmlBody, final_date, (", ".join(open_men)), messaged, event_location)
        s = ""
        print("|--- End: " + customer_name + " ---|")
        print(" ")


