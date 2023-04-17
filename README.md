# Google Mail Merge App

This project was inspired by a need in our local district to create a more flexible way to send out mass communications securely. 

### Requirements 
- Picker V1
- Drive V2
- GCP (Integrations and Org Access)
- API Key for Drive Authentication

### OAuth Scopes
- https://www.googleapis.com/auth/userinfo.email
- https://www.googleapis.com/auth/userinfo.profile
- https://mail.google.com/
- https://www.googleapis.com/auth/documents
- https://www.googleapis.com/auth/drive
- https://www.googleapis.com/auth/forms
- https://www.googleapis.com/auth/presentations
- https://www.googleapis.com/auth/script.container.ui
- https://www.googleapis.com/auth/spreadsheets


## Summary

This application is a three step process. The first step inspects your data with a few prompts and will make sure the neccessary items are included. If everything is correct you will get a confirmation message. You will then either select a document to use as a template or be provided with one. This is primarilly to give you an example on how to use the tags that will refer to the data in your spreadsheet. Finally, the last step will create your documents and give you the option to email the recipients in the email column. 
        
This app creates a folder in your google drive called "Mail Merge Drafts" and at the time of creating the template documents creates a subfolder to store all the templates. Everything can be accessed in this folder at a later date if needed.
        
You will notice a "Confirmation" sheet added to your file. This contains links to each individual file and a Yes/No indicating wether an email was sent.
