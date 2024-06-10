import requests, json, getpass

resource = "/parse/login"
 
def login(userName, serverurl, AppID):
    password = 'Loldosloll@123' 
    header = {'Content-Type': 'application/json',
          'X-Parse-Application-Id': AppID} 

    payload = {'username': userName, 
           'password': password}
    response_decoded_json = requests.post(serverurl + resource, json=payload, headers=header)
    response_dict = response_decoded_json.json()
    return(response_dict['sessionToken'])


