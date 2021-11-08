import requests
import csv
import sys
import json
import logging
import msal
import os
import atexit
import pyAesCrypt
from new_hire import NewHire
from datetime import date


# This script pulls data from a CSV then uses Microsoft Graph API to automate sending of welcome emails to new hires.
# This script runs daily to check for matching new hires, send each a welcome email, then send a report email to the
# Site IT Dev team.
# Microsoft Graph is a RESTful web API that enables you to access Microsoft Cloud service resources. After you register 
# your app and get authentication tokens for a user or service, you can make requests to the Microsoft Graph API.
# To learn more about Microsoft Graph API, please visit https://docs.microsoft.com/en-us/graph/use-the-api
# To explore Microsoft Graph API's toolsets, please visit https://developer.microsoft.com/en-us/graph/graph-explorer


# pw = ''                                                           # Use only for testing
# today = '2022-09-27'                                              # Use only for testing
today = date.today().strftime("%Y-%m-%d")                           # New hire's start date
site_location = 'Redwood City'                                      # New hire's location
csv_url = ""                                                        # Csv link managed by...
pw = os.environ['DEV_PASS']                                         # A password must first be set in the environmental variables
cache_file = r'C:\Projects\new_hire_welcome\my_cache.bin'           # Cache file path which stores Access Token value

tenant_id = os.environ['TENANT_ID']                                 # Tenant value found in Azure AD
app_id = os.environ['APP_ID']                                       # App ID found in Azure AD App Registration
secret = os.environ['SECRET']                                       # Client secret found in Certificates and Secrets right after App Registration

app_config = {
    "client_id":f"{app_id}",
    "authority":f"https://login.microsoftonline.com/{tenant_id}",
    "client_secret":f"{secret}",
    "scope":["https://graph.microsoft.com/Mail.Send",               # Resource scope needed so that permissions needed is known
             "https://graph.microsoft.com/User.Read"],
    "endpoint":"https://graph.microsoft.com/v1.0/me/sendMail"       # Endpoint needed to send emails via API
}

def encrypt_cache(passw):
    output = cache_file + '.aes'                                    # Creates output file name
    try:
        pyAesCrypt.encryptFile(cache_file, output, passw)           # Encrypts the cache file and outputs a new encrypted cache file
        os.remove(cache_file)                                       # Deletes the original unecrypted cache file
    except Exception as e:
        print(e)                                                    # Returns and Exception if there is an error

def decrypt_cache(passw):
    incoming = cache_file + '.aes'                                  # Sets the input as the encrypted cache file
    try:
        pyAesCrypt.decryptFile(incoming, cache_file, passw)         # Decrypts the encrypted cache file and outputs a new unencrypted cache file
        os.remove(incoming)                                         # Deletes the encrypted cache file
    except Exception as e:
        print(e)                                                    # Returns and Exception if there is an error

def csv_parser(url):                                                # Date must be in "%Y-%m-%d" format, this returns a new list of New Hire objects
    response = requests.get(url)                                    # Pulls all data from the csv url
    rows = response.text.split('\r\n')                              # Divides the data by lines and creates a list
    dct = csv.DictReader(rows)                                      # Sets the first element of the list as the header and the following elements as the value. It also converts the structure into a list of dictionaries.

    new_list = []                                                   # Initialize a new list of dictionary
    for each_item in dct:                                           # Iterates throught the dict list to check for conditions
        if each_item['Start_Date'] == today:                        # Checks if Start_Date value matches today
            if each_item['Location'] == "Redwood City":             # Checks if Location value matches Redwood City
                if len(each_item['Email']) != 0:                    # Checks if length of Email value is not 0 (or not empty)
                    new_list.append(                                # Creates a NewHire instance of the matching values and adds it to new_list
                        NewHire(each_item['Name'], each_item['Start_Date'], each_item['Location'], each_item['Email'])
                    )
    return new_list                                                 # Returns new_list once iteration has finished going through entire dict list

def get_and_cache_token(config):
    cache = msal.SerializableTokenCache()                           # Creates an instance of a cache object
    if os.path.exists(cache_file):                                  # Checks if cache_file already exist
        print("token cache exists")                                 # prints cache exists if cache_file exists
        cache.deserialize(open(cache_file, "r").read())             # opens cache_file and reads content
    atexit.register(
        lambda: open(cache_file, "w").write(cache.serialize())
        # Hint: The following optional line persists only when state changed
        if cache.has_state_changed else None
        )

    # Create a preferably long-lived app instance which maintains a token cache.
    app = msal.PublicClientApplication(
        config["client_id"], authority=config["authority"],
        token_cache=cache                                           # Default cache is in memory only.
    )                                                               # You can learn how to use SerializableTokenCache from
                                                                    # https://msal-python.readthedocs.io/en/latest/#msal.SerializableTokenCache
                                                                    # The pattern to acquire a token looks like this.
    result = None
                                                                    # Note: If your device-flow app does not have any interactive ability, you can
                                                                    # completely skip the following cache part. But here we demonstrate it anyway.
                                                                    # We now check the cache to see if we have some end users signed in before.
    accounts = app.get_accounts()
    if accounts:
        logging.info("Account(s) exists in cache, probably with token too. Let's try.")
        print("Pick the account you want to use to proceed:")
        for a in accounts:
            print(a["username"])
                                                                    # Assuming the end user chose this one
        chosen = accounts[0]
                                                                    # Now let's try to find a token in cache for this account
        result = app.acquire_token_silent(config["scope"], account=chosen)

    if not result:
        logging.info("No suitable token exists in cache. Let's get a new one from AAD.")

        flow = app.initiate_device_flow(scopes=config["scope"])
        if "user_code" not in flow:
            raise ValueError(
                "Fail to create device flow. Err: %s" % json.dumps(flow, indent=4))

        print(flow["message"])
        sys.stdout.flush()                                          # Some terminal needs this to ensure the message is shown

                                                                    # Ideally you should wait here, in order to save some unnecessary polling
        # input("Press Enter after signing in from another device to proceed, CTRL+C to abort.")

        result = app.acquire_token_by_device_flow(flow)             # By default it will block
                                                                    # You can follow this instruction to shorten the block time
                                                                    #    https://msal-python.readthedocs.io/en/latest/#msal.PublicClientApplication.acquire_token_by_device_flow
                                                                    # or you may even turn off the blocking behavior,
                                                                    # and then keep calling acquire_token_by_device_flow(flow) in your own customized loop.

    if "access_token" in result:
        return result['access_token']                               # Returns the Access Token value if it is in result

    else:
        print(result.get("error"))
        print(result.get("error_description"))
        print(result.get("correlation_id"))                         # You may need this when reporting a bug

def send_welcome_email(token, config, new_emp):

                                                                    # Calling API using the access token
    response = requests.post(                                       # Use token to call downstream service
                config["endpoint"],                                 # Gives the endpoint URL
                headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'},
                json=new_emp.getpayload()                           # Uses the payload created by New_Hire object
    )

    return f'''Status Code: {str(response.status_code)}             
            Response: {str(response)} 
            Name: {new_emp.get_fullname()} 
            Email: {new_emp.get_email()}'''                         # Returns these values for error reporting

def send_report_email(token, config, content):
    distro = 'ears-siteit-dev@ea.com'                               # This is the distro where error reporting is sent to.
    payload = {                                                     # This creates the payload structure that is sent with error reporting
        "message": {
            "subject": 'Welcome New Hire Report',
            "body": {
                "contentType": 'HTML',
                "content": content,                                 # Status Code, Response, Name, and Email will be passed as the content
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": distro
                    }
                }
            ]
        }
    }

    requests.post(                                                  # Use token to send report via Graph API
        config["endpoint"],
        headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'},
        json=payload
    )

rws_newhires = csv_parser(csv_url)                                  # Pulls csv data and converts to list of New Hire objects and passes the list to rws_newhire variable
decrypt_cache(pw)                                                   # Decrypts my_cache.bin.aes --> my_cache.bin
token = get_and_cache_token(app_config)                             # Gets token
for each in rws_newhires:                                           # Iterates through list of New Hire objects
    print(each.get_fullname(), each.get_email())                    # Prints the full name and email of the New Hire object
    response_data = send_welcome_email(token, app_config, each)     # Sends an email to new hire using the token, app configuration, and New Hire object and passes status code to response_data variable
    send_report_email(token, app_config, response_data)             # Sends response data as report to the distro in send_report_email
encrypt_cache(pw)                                                   # Encrypts my_cache.bin --> my_cache.bin.aes
