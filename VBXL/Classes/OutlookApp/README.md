# OutlookApp

A late bound object made to utilize Outlook functionalities from Excel.
- Uses the Outlook application on the current user's machine.
    - The `Outlook.Application` object.
- Uses the accounts on the user's Outlook application.

## Properties

| Property         | Type      | Description                                                                               |
|------------------|-----------|-------------------------------------------------------------------------------------------|
| `IsInstalled`    | `Boolean` | If Outlook is installed (read-only).                                                      |
| `EmailAddress`   | `String`  | The email address to use.                                                                 |
| `Account`        | `Object`  | The email account object (read-only). Set by `EmailAddress`.                              |
| `EmailSignature` | `String`  | The email account's default signature for new emails (read-only).  Set by `EmailAddress`. |
| `UseSignature`   | `Boolean` | Whether or not to use the signature in the email body.                                    |

## Methods & Functions

|                                 | Description                                                                   |
|---------------------------------|-------------------------------------------------------------------------------|
| [`GetAccount`](#getaccount)     | Gets a single Outlook account that's added on the user's Outlook application. |
| [`GetSignature`](#getsignature) | Gets the default email signature for a new email from an account.             |
| [`CreateDraft`](#createdraft)   | Creates an Outlook draft.                                                     |

---


### [`GetAccount`](OutlookApp.cls#L183)

Gets a single Outlook account that's added on the user's Outlook application.

**Parameters**
- `olEmailFrom` `ByRef` `Variant`
    - The email address or the Outlook account.

**Returns**
- `Object`: The Outlook account.


### [`GetSignature`](OutlookApp.cls#L203)

Gets the default email signature for a new email from an account.
- Has to create and display a draft to extract the signature from the draft's HTML body.
- Images are not imported correctly due to the <img>'s src so they are removed.

**Parameters**
- `olEmailFrom` `ByRef` `Variant`
    - The email address or the Outlook account.

**Returns**
- `String`: The account's email signature as HTML.



### [`CreateDraft`](OutlookApp.cls#L296)

Creates an Outlook draft.

**Parameters**
- `olEmailBody` `ByVal` `String`
    - Content for the HTML email body. Could just be text provided.
- `olSubject` `ByVal` `String`
    - The email subject.
- `olEmailTo` `ByVal` `String`
    - The email TO recipients.
    - Each recipients is delimited by a semi-colon (;).
- `olEmailCc` `ByVal` `String`
    - The email CC recipients.
    - Each recipients is delimited by a semi-colon (;).
- `olAttachments` `ByRef` `Variant` [`Optional`]
    - An array of the attachments to include.
    - Each item in the array is a full file path.

---

## Usage

```vb
Private Sub Demo()
    Dim App                 As New OutlookApp
    ReDim Attachments(1)    As Variant

    ' The attachments have to consist of a FULL file path
    Attachments(0) = "Some file to attach1.jpg"
    Attachments(1) = "Some file to attach1.jpg"

    App.EmailAddress = "YourEmailAddress@domain.com"
    App.CreateDraft "Hello World", "TEST", "Recipient1@domain.com; Recipient2@domain.com", olAttachments:=Attachments

    ' Sending an email without a signature    
    App.UseSignature = False
    App.CreateDraft "Hello World", "TEST", "Recipient1@domain.com; Recipient2@domain.com", olAttachments:=Attachments
End Sub
```