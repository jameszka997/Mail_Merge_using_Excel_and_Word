# Mail_Merge_using_Excel_and_Word
This repository aims to create an easy to understand guide to creating mail merge list and e-mail template for tailored and customized e-mailing for marketing and information sharing.

---

## Purposes of this repository
- Documenting the manner in which First Name, Last Name and e-mail was extracted and formatted for easy use from merged strings using Excel functions (XLOOKUP, SUMPRODUCT)
- Setting up Mail Merge within Word for easy to use customized e-mail sharing to many participants.
- Workflow for assembling project specific participants and how relevant participants were cross-checked and filtered from all participants list.


***Disclaimer: This project was part of a company project which involved sensitive information, thus any data and information used during the project is fabricated.***

## Assembling the Excel file for the Mail Merge
Originally I have had Multiple Projects within our database for which I have needed to look up all stakeholders that are connected to them.
After compiling all of them I have a dataset which I can work from.

Project	Stakeholder
Project A	Company A
Project A	Company B
Project A	Company C
Project A	Company D
Project A	Company E
Project B	Company A
Project B	Company B
Project B	Company E
Project C	Company F
Project C	Company H
Project C	Company I
Project C	Company A
Project C	Company O
Project C	Company P
Project C	Company Q
Project C	Company S
Project C	Company T
Project D	Company A
<img width="145" height="381" alt="image" src="https://github.com/user-attachments/assets/977cb5a9-eb6c-463e-b861-22ea73c1fbd9" />

Some of these stakeholders are legacy entries which are no longer functioning, thus I needed to cross-check from the list of valid stakeholders and remove any unnecessary ones from the list.

| Valid Stakeholders |
| -- |
|Company A|
|Company C|
|Company D|
|Company F|
|Company H|
|Company I|
|Company O|
|Company P|
|Company Q|


`SUMPRODUCT()` excel function was used to create a supporting column which returns *TRUE* value if the cell entry is within the pre-defined list, Valid Stakeholders column in our case, and 'FALSE' if it is not part of it.

`=SUMPRODUCT(--ISNUMBER(SEARCH($I$2:$I$10, B2)))>0`

### Output
Project	Stakeholder	Stakeholder 
Project A	Company A	TRUE
Project A	Company B	FALSE
Project A	Company C	TRUE
Project A	Company D	TRUE
Project A	Company E	FALSE
Project B	Company A	TRUE
Project B	Company B	FALSE
Project B	Company E	FALSE
Project C	Company F	TRUE
Project C	Company H	TRUE
Project C	Company I	TRUE
Project C	Company A	TRUE
Project C	Company O	TRUE
Project C	Company P	TRUE
Project C	Company Q	TRUE
Project C	Company S	FALSE
Project C	Company T	FALSE
Project D	Company A	TRUE
<img width="209" height="381" alt="image" src="https://github.com/user-attachments/assets/cdc14045-3bb6-48de-baa4-49d47d14dea6" />





## Resources
The following resources were of chief importance in assembling this project:
- [How to Mail Merge in Microsoft Word](https://www.wikihow.com/Mail-Merge-in-Microsoft-Word)
- [Word: Mail merge by LearnFree](https://www.youtube.com/watch?v=do9ujnZLIC4&embeds_referring_euri=https%3A%2F%2Fwww.bing.com%2F&embeds_referring_origin=https%3A%2F%2Fwww.bing.com&source_ve_path=MjM4NTE)


