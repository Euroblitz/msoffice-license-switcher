# msoffice-license-switcher
![Screenshot](blob:https://github.com/be361ac6-f2de-4e1a-8493-fa2b6087b3eb)
Batch scripts made in .vbs to easily add and remove Microsoft Office license keys. Useful for testing a big number of keys bought by a reseller. I had to test almost 200 Office 2021 keys by hand because we had a few issues with some bought by a reseller, so I did a script to test them easily.



## For adding a key
Usage: Run "keyAdd.vbs" (it will ask for adming privileges), paste the key and confirm, and open Office to check if the activation works.

## For removing a key
Usage: Run "keyRemove.vbs" (it will ask for adming privileges), paste the key and delete everything else except the last 5 digits and confirm, and open Office to check if the activation works.
