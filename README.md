# PptPolly - Add Amazon Polly voice to a PowerPoint presentation

Yet another spare time piece of code to automate things I do...

## Instructions

- Get code
- Download and install Visual Studio (I used <a href="https://visualstudio.microsoft.com/downloads/" target="_blank">Community Edition</a>)
- Configure the computer to develop Office solutions (<a href="https://docs.microsoft.com/en-us/visualstudio/vsto/configuring-a-computer-to-develop-office-solutions?view=vs-2017" target="_blank">link</a>)
- You may need to install the Nugets for AWSSDK.Core and AWSSDK.Polly (hopefully downloaded automatically, or follow <a href="https://chrisbitting.com/2017/04/07/using-amazon-polly-from-net-c-get-mp3-file/" target="_blank">link</a>)
- After opening the PptPolly.sln, open app.config and set the values for your AWS AccessKeyID and SecretKey. Make sure you set the policy AmazonPollyFullAccess for the <a href="https://console.aws.amazon.com/iam/home?#" target="_blank">IAM user</a>.
- Compile and start the project. Go to a PowerPoint slide and enter these lines in the notes
```
This is a test
Matthew
```
- Go to Add-ins tab, and in the PptPolly group, click "Add Voice"
- After a brief webservice call, and audio should be inserted in the slide
- If there is a textbox with title "Caption", the first line appears in the textbox

## TODOs
- Investigate why audio won't playback automatically. But if go to Animation and change to something else and then back to automatic, it will playback.
- Progress during webservice call?
- Button to automate task for all slides at once
- Button to remove all other audio (Slide Show -> Record Slide Show -> Clear Narration on All Slides)

# References

Visual Studio Downloads
https://visualstudio.microsoft.com/downloads/

Using Amazon Polly from .NET / C#, Get MP3 file
https://chrisbitting.com/2017/04/07/using-amazon-polly-from-net-c-get-mp3-file/

Understanding and Getting Your Security Credentials
https://docs.aws.amazon.com/general/latest/gr/aws-sec-cred-types.html#access-keys-and-secret-access-keys

Walkthrough: Create your first VSTO Add-in for PowerPoint
https://docs.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-powerpoint?view=vs-2017

Powerpoint add-on to get text in notes in slides and convert it to audio. Doesn't seem to be getting the notes in the slides like it should?
https://stackoverflow.com/questions/20975165/powerpoint-add-on-to-get-text-in-notes-in-slides-and-convert-it-to-audio-doesn

