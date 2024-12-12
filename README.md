# Cleantalk VBScript Example

This project demonstrates how to use the Cleantalk API with VBScript to validate registrations.

## Requirements

   1. Register a Cleantalk account https://cleantalk.org/register?product=anti-spam

   2. Obtain the access key from the CleanTalk account https://cleantalk.org/help/add-website

## Usage

1. **Place the code of [Cleantalk class](https://github.com/alexandergull/cleantalk-vbs/blob/master/cleantalk_example.vbs#L3) in your VBScript file.**
    
        ```vbscript
        Class CleantalkClass
            Private auth_key
            Private check_message
            Private user_email
            Private user_ip
            Private user_name
            Private user_js_state
            Private user_submit_time
            Private form_event_token
            Private response
            Private verdict
            Private codes
            Private comment
            ... other class code
            ... other class code
          else
               validateResponse = false
          end if
            end function
         
         end class
        ```

2. **Initialize the Cleantalk class instance, use your own access key when instantiating:**

    ```vbscript
    Dim Cleantalk : Set Cleantalk = (New CleantalkClass)("your_access_key", "check_newuser")
    ```

3. **Set user data when the logic is ready to check the user:**

    ```vbscript
    Cleantalk.setUserEmail("stop_email@example.com")
    Cleantalk.setUserIP("10.10.10.10")
    Cleantalk.setUserName("John Doe")
    Cleantalk.setUserJSState("1")
    Cleantalk.setUserSubmitTime("5")
    Cleantalk.setFormEventToken("a_32_symbols_event_token_value")
    ```

4. **Send request to the API:**

    ```vbscript
    Cleantalk.sendRequest
    ```

5. **Validate the response:**

    ```vbscript
    If Cleantalk.validateResponse Then
        If Cleantalk.getVerdict = 1 Then
            WScript.Echo "Validation success. User is allowed."
        Else
            WScript.Echo "Validation success. User is blocked. Reason: " & Cleantalk.getCodes & " " & Cleantalk.getComment
        End If
    Else
        WScript.Echo "Validation failed. Code: " & Cleantalk.getCodes & " Comment: " & Cleantalk.getComment
    End If
    ```
   **Important!** Do validation every time after response gathering.

## Functions

- **setUserEmail(email)**
- **setUserIP(ip)**
- **setUserName(name)**
- **setUserJSState(jsState)**
- **setUserSubmitTime(submitTime)**
- **setFormEventToken(eventToken)**
- **sendRequest()**
- **validateResponse()**
- **getVerdict()**
- **getCodes()**
- **getComment()**

## Example

```vbscript
Dim Cleantalk : Set Cleantalk = (New CleantalkClass)("your_auth_key", "check_message")

Cleantalk.setUserEmail("user@example.com")
Cleantalk.setUserIP("192.168.1.1")
Cleantalk.setUserName("John Doe")
Cleantalk.setUserJSState("0")
Cleantalk.setUserSubmitTime("0")
Cleantalk.setFormEventToken("your_event_token")

Cleantalk.sendRequest

If Cleantalk.validateResponse Then
    If Cleantalk.getVerdict = 1 Then
        WScript.Echo "Validation success. User is allowed."
    Else
        WScript.Echo "Validation success. User is blocked. Reason: " & Cleantalk.getCodes & " " & Cleantalk.getComment
    End If
Else
    WScript.Echo "Validation failed. Code: " & Cleantalk.getCodes & " Comment: " & Cleantalk.getComment
End If
```

## Implementing of BotDetector JavaScript library

To use the BotDetector JavaScript library, you need to include the script in the HTML of the page.

```html
<script src="https://moderate.cleantalk.org/ct-bot-detector-wrapper.js"></script>
```

This script will automatically detect the form submission event and send the data to the Cleantalk API.

Please note, that the script does not perform any checks, just sends the user's frontend data (like JavaScirpt state, mouse position etc.) to the API. 

### Example
```html
<!DOCTYPE html>
<html lang="en">
   <head>
   <meta charset="UTF-8">
   <title>Register</title>
   <!--Bot-detector JS library wrapper. This script must be added to the HTML of the page.-->
   <script src="https://moderate.cleantalk.org/ct-bot-detector-wrapper.js"></script>
</head>
<body>
   <form method="post" action="your_form_handler_script">
      <label for="user_name">User name</label>
      <label for="user_email">User email</label>
      <input type="text" name="user_name" id="search_field" /> <br />
      <input type="text" name="user_email" id="search_field" /> <br />
      <input type="submit" />
   </form>
</body>
</html>
```

When you got added the script, the form will be updated with hidden event_token field after the script loaded. This field value you should transfer to [VB Script](https://github.com/alexandergull/cleantalk-vbs/blob/master/cleantalk_example.vbs#L184)

Once the token is provided in the API request, the VB Script will make the API takes in count the frontend data.

Make note, that data provided on event_token have the higher priority than the data set by optional VB Script setters.
