10/16/24
Implemented a while loop to attempt to combine both ExchangeOnline and Microsoft Graph into one script.
Currently, working on getting the authentication to not spam when selecting option 1.

10/17/24
Hard coded domain and user to get around the authentication spam when selecting 1.
Added try/catch block to catch error exceptions.
Added delay to allow script to run properly.
Fixed up Foreach block in selection 1.
Added Invalid character checker between lines 12-14

10/18/24
Made some edits to the list of Options.
Added a block for third selection.
Removed Two Write-host from 1 selection block "Check Cred" & "Login Successful".
Added a new module for revoke sessions support.
Added two blocks for mailbox conversion of both shared to regular, vice versa.

10/23/24
Added another elseif option for revoking sessions for #5
Included "User.ReadWrite.All" for manipulating user sessions as a scope found within connect-mggraph
Currently trying to implement disabling devices but found that it doesnt pipe with get-mguser. At a stand still at this part..

10/24/24
Added confirmation in the License block, asking if the user wants to delete found licenses with Y/N.
Added more delays between code blocks with Write-Host "Returning back to home".