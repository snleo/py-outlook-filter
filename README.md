# py-outlook-filter
Outlook junk email filter using Python

## Packages required
- pywin32
- re
- yaml

You can use `pip` to install all packages

```pip install pypiwin32```

```pip install re```

```pip install pyyaml```

## settings.yaml
All settings can be found in this file

```
keywords_file : "junkwords.txt"
outlook_account : "Website"
outlook_inbox : "Inbox"
outlook_junk : "Junk E-Mail"
```

## Keyword file
If Subject or Body of a message contains any word in the spcified keyword file it will be moved to specified Outlook folder.

That's it! Just run the script to filter outlook mailbox!
