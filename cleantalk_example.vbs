option explicit

class CleantalkClass
     Dim xmlhttp

     private agent
     private moderate_url
     private auth_key
     private method_name
     private user_email
     private user_ip
     private user_name
     private js_on
     private submit_time
     private request_json
     private event_token
     private response_json
     private verdict_allowed
     private verdict_account_status
     private verdict_codes
     private verdict_comment
     private verdict_allowed_pattern
     private verdict_account_status_pattern
     private verdict_codes_pattern
     private verdict_comment_pattern

     ' Constructor, set defaults
     Private Sub Class_Initialize()
          agent = "vbscript-1.0"
          moderate_url = "https://moderate.cleantalk.org/api2.0"
          auth_key=""
          user_email=""
          user_ip=""
          user_name=""
          method_name="check_newuser"
          js_on=1
          submit_time=5
     End Sub

     ' default constructor on init
     Public default function init(pAPIkey, pMethodName)
          auth_key = pAPIkey
          method_name = pMethodName

          Set Init = Me
     end function

     ' Setters

     public Sub setUserEmail(pEmail)
          user_email=pEmail
     End Sub

     public Sub setUserIP(pIP)
          user_ip=pIP
     End Sub

     public Sub setUserName(pName)
          user_name=pName
     End Sub

     public Sub setUserJSState(pJSState)
          js_on=pJSState
     End Sub

     public Sub setUserSubmitTime(pSubmitTime)
          submit_time=pSubmitTime
     End Sub

     public Sub setFormEventToken(pFormEventToken)
          event_token=pFormEventToken
     End Sub

     ' Getters

     public function getVerdict()
          getVerdict = verdict_allowed
     end function

     public function getAccountStatus()
          getAccountStatus = verdict_account_status
     end function

     public function getCodes()
          getCodes = verdict_codes
     end function

     public function getComment()
          getComment = verdict_comment
     end function

     ' Construct JSON data

     private sub constructJSONData()
          request_json = "{""auth_key"":""" & auth_key & """,""agent"":""" & agent & """,""method_name"":""" & method_name & """,""sender_email"":""" & user_email & """,""sender_ip"":""" & user_ip & """,""sender_nickname"":""" & user_name & """,""js_on"":""" & js_on & """,""submit_time"":""" & submit_time & """, ""event_token"":""" & event_token & """}"
     end sub

     ' Send request to API

     public sub sendRequest()
          constructJSONData
          ' Create an XMLHTTP object
          Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

          ' Open a connection to the API
          xmlhttp.open "POST", moderate_url, False

          ' Set the request headers
          xmlhttp.setRequestHeader "Content-Type", "application/json"

          ' Send the request
          xmlhttp.send request_json

          ' Get the response
          response_json = xmlhttp.responseText

          ' Clean up
          Set xmlhttp = Nothing
     end sub

     ' Parse JSON response
     private sub ParseJSONResponse()
          verdict_allowed_pattern = """allow"":([\d]),"
          verdict_account_status_pattern = """account_status"":([\d]),"
          verdict_codes_pattern = """codes"":""(.*)"""
          verdict_comment_pattern = """comment"":""(.*)"""
          verdict_allowed = regexSearch(response_json, verdict_allowed_pattern)
          verdict_account_status = regexSearch(response_json, verdict_account_status_pattern)
          verdict_codes = regexSearch(response_json, verdict_codes_pattern)
          verdict_comment = regexSearch(response_json, verdict_comment_pattern)
     End sub

     ' Search for a pattern in a string
     private function regexSearch(pJSONResponse, pPattern)
          Dim regEx, matches

          ' Create a RegExp object
          Set regEx = New RegExp
          regEx.Pattern = pPattern
          regEx.IgnoreCase = True
          regEx.Global = True

          ' Execute the search
          Set matches = regEx.Execute(pJSONResponse)
          if matches.Count = 0 then
               regexSearch = ""
               exit function
          end if
          if matches(0).Submatches.Count = 0 then
               regexSearch = ""
               exit function
          end if
          regexSearch = matches(0).Submatches(0)
     end function

     ' Validate the response
     public function validateResponse()
          ParseJSONResponse
          wscript.echo "Response: " & response_json
          if getAccountStatus = 1 then
               validateResponse = true
          else
               validateResponse = false
          end if
     end function

end class

' Usage example
' Init a Cleantalk class instance
dim Cleantalk : Set Cleantalk = (New CleantalkClass)("your_access_key", "check_newuser")

' there can be placed your site code
' SOME LOGIC BEFORE CLEANTALK CHECK, like getting user data from form

' Init user data when the properties is ready
Cleantalk.setUserEmail("stop_email@example.com") ' Required, if unused or empty - empty value will be used
Cleantalk.setUserIP("10.10.10.10") ' Required, if unused or empty - empty value will be used
Cleantalk.setUserName("John Doe") ' Optional, if unused or empty - empty value will be used
Cleantalk.setUserJSState("1") ' Optional, if unused or empty - 1 will be used
Cleantalk.setUserSubmitTime("5") ' Optional, if unused or empty - 5 will be used

' Set form event token, you would attach a JS script to the page with form to use this, read more in README.md
Cleantalk.setFormEventToken("a_32_symbols_event_token_value") ' Optional, if empty - empty value will be used

' Send request to API
Cleantalk.sendRequest

' Validate the response, always run this before getting the verdict
if Cleantalk.validateResponse then
   if Cleantalk.getVerdict = 1 then
     'do something if user is allowed
     wscript.echo "Validation succes. User is allowed."
   else
     'do something if user is blocked
     wscript.echo "Validation succes. User is blocked. Reason: " & Cleantalk.getCodes & " " & Cleantalk.getComment
   end if
else
     'do something if validation failed
     wscript.echo "Validation failed. Code: " & Cleantalk.getCodes & " Comment: " & Cleantalk.getComment
end if

' Clean up
Set Cleantalk = Nothing
