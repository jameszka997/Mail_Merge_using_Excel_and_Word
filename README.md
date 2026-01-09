<img width="97" height="88" alt="image" src="https://github.com/user-attachments/assets/226fc281-9b3e-49ed-99b5-222e66ac28f8" /># Mail_Merge_using_Excel_and_Word
This repository aims to create an easy to understand guide to creating mail merge list and e-mail template for tailored and customized e-mailing for marketing and information sharing.

---

## Purposes of this repository
- Documenting the manner in which First Name, Last Name and e-mail was extracted and formatted for easy use from merged strings using Excel functions (XLOOKUP, SUMPRODUCT)
- Setting up Mail Merge within Word for easy to use customized e-mail sharing to many participants.
- Workflow for assembling project specific participants and how relevant participants were cross-checked and filtered from all participants list.


***Disclaimer: This project was part of a company project which involved sensitive information, thus any data and information used during the project is fabricated.***


## Structure of the Excel File for the Mail Merge Process

| Project   | Stakeholder | Last Name | First Name | E-mail                          |
|-----------|-------------|-----------|------------|--------------------------------|
| Project A | Company A   | Nagy      | Anna       | anna.nagy@companya.com         |




As a start, we will start by opening a *blank* Word page for this tutorial, and finding the **Mailings** option on the Ribbon.
<img width="954" height="275" alt="Word 1" src="https://github.com/user-attachments/assets/aa8bcfd7-0a62-4148-a461-1df818f0378d" />

Then we click on the **Start Mail Merge** option.
![Word 2](https://github.com/user-attachments/assets/6f89f9ed-b7d4-47ab-813e-11e19cf121c7)

Then we click on the **Step-by-step Mail Merge Wizard** option. This will start the mail merge process which can be followed.
![Word 3](https://github.com/user-attachments/assets/398fb2a3-0a70-4b88-8ae0-53cf02fc8cb5)

I have prepared the primary text for the e-mail, but this can be performed at any step of the way.
![Word 4](https://github.com/user-attachments/assets/fe79b85a-5d8f-4ef7-ad0c-2e12414adbae)


`Step 1 - Select Document Type`. The E-mail template is chosen from these, but depending on the purpose of the document (it is intended for printing, or other purpose) additional ones can be chosen.
![Word 5](https://github.com/user-attachments/assets/a518a0b7-af8a-4c57-8464-a10899857475)

`Step 2 - Select starting document`. Existing documents or templates can also be utilized, for this guide, the *Use current document* option is chosen.
![Word 6](https://github.com/user-attachments/assets/a1b48f5b-5352-44f1-a304-af77745a4fee)

'Step 3 - Select recipients'. Existing outlook contacts can be chosen, or even typed out from scratch, we shall choose *Use an existing list* and browsing for the excel list already prepared.
![Word 7](https://github.com/user-attachments/assets/1197ca8e-2c27-42a1-bf08-86a98a4be6f1)

Once chosen, a new window will pop up with the sheets available within the Excel file, we would choose the one holding our recipient table.
![Word 8](https://github.com/user-attachments/assets/c0b90d4c-ace7-4cd3-98b4-3ace127258af)

Then a new window will pop-up which showcases our table's columns and information within it. The table's information can be freely viewed, sorted, filtered and edited here if chosen so.
![Word 9](https://github.com/user-attachments/assets/8fa32e1f-9b25-4a91-8cea-0359200e6747)

'Step 4 - Write your e-mail message' relates to adding the sections to the e-mail template which would add the information intended for each individual e-mail, such as the name of the recipient, and other information intended for each e-mail and defined within the Excel file. As a beginning, I shall add the *Greeting line* to the e-mail which will start the e-mail.
![Word 10](https://github.com/user-attachments/assets/63c56259-e7d7-4bcb-85c1-878d879d66b3)

A new window will pop-up where the details of the greeting line can be edited. The default starts with `Dear First Name Last Name,`, as Word automatically tries to find First and Last Name columns from the Excel but can be reassigned with the `Match Fields...` option. Each recipient greeting can be previewed in the picture. Example from the picture `Dear Anna Nagy,`
![Word 11](https://github.com/user-attachments/assets/76799f2f-f045-480e-8dda-e5d17107a848)

Once Okay is pressed, the Greeting line shall appear like this, highlighted on the picture.
![Word 12](https://github.com/user-attachments/assets/d2a381e9-ffc4-4889-ba5d-d5d5b7d5809b)

Additional information can be added via `More items...` option *see picture below*. This shall be used to add the `Stakeholder` and `Project` information to the e-mail for making it more personal. 
![Word 13](https://github.com/user-attachments/assets/367c3211-42b0-4aec-a410-9d0c21889980)

In the pop-up window, just select the appropriate column from the Excel and insert it where it is needed in the e-mail. The text of the e-mail would look like this.
![Word 14](https://github.com/user-attachments/assets/b99f7dee-0559-45fd-8491-4cff795b836c)

'Step 5 - Preview your e-mail messages' this allows each recipient's e-mail to be viewed, find a recipient or excluse and edit the recipient list. Example is given below how one of these customized e-mail would look like.
![Word 15](https://github.com/user-attachments/assets/2de90f25-f969-46d7-a0e7-2741c90aac2d)

'Step 6 - Mail Merge', this step would send out all the e-mails once clicked on the `Electronic Mail...` button. The e-mail would be sent from the Outlook e-mail logged in from the device.
![Word 17](https://github.com/user-attachments/assets/0e5adcf1-68da-4642-b265-6f4d8ceed007)









## Appendix - Assembling the Excel file for the Mail Merge
Originally I have had Multiple Projects within our database for which I have needed to look up all stakeholders that are connected to them.
After compiling all of them I have a dataset which I can work from.



| Project   | Stakeholder |
|-----------|-------------|
| Project A | Company A   |
| Project A | Company B   |
| Project A | Company C   |
| Project A | Company D   |
| Project A | Company E   |
| Project B | Company A   |
| Project B | Company B   |
| Project B | Company E   |
| Project C | Company F   |
| Project C | Company H   |
| Project C | Company I   |
| Project C | Company A   |
| Project C | Company O   |
| Project C | Company P   |
| Project C | Company Q   |
| Project C | Company S   |
| Project C | Company T   |
| Project D | Company A   |



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

| Project   | Stakeholder | Status |
|-----------|-------------|--------|
| Project A | Company A   | TRUE   |
| Project A | Company B   | FALSE  |
| Project A | Company C   | TRUE   |
| Project A | Company D   | TRUE   |
| Project A | Company E   | FALSE  |
| Project B | Company A   | TRUE   |
| Project B | Company B   | FALSE  |
| Project B | Company E   | FALSE  |
| Project C | Company F   | TRUE   |
| Project C | Company H   | TRUE   |
| Project C | Company I   | TRUE   |
| Project C | Company A   | TRUE   |
| Project C | Company O   | TRUE   |
| Project C | Company P   | TRUE   |
| Project C | Company Q   | TRUE   |
| Project C | Company S   | FALSE  |
| Project C | Company T   | FALSE  |
| Project D | Company A   | TRUE   |


Then I have used an `XLOOKUP()` function which checked what Stakeholder is from my reference list (Column 1), and then put the corresponding contact information (Column 2) into the cell so I would have a contact column with the appropriate contact information for each company.

`=XLOOKUP(TRUE, ISNUMBER(SEARCH($I$2:$I$26, D10)), $J$2:$J$26, "No match")`

## Output


| Project   | Stakeholder | Contact                                   |
|-----------|-------------|-------------------------------------------|
| Project A | Company A   | Nagy, Anna /AA <anna.nagy@companya.com>   |
| Project A | Company B   | Smith, John /AA <john.smith@companyb.com> |
| Project A | Company C   | Kovacs, Peter /AA <peter.kovacs@companyc.com> |
| Project A | Company D   | Garcia, Maria /AA <maria.garcia@companyd.com> |
| Project A | Company E   | Rossi, Luca /AA <luca.rossi@companye.com> |
| Project B | Company A   | Nagy, Anna /AA <anna.nagy@companya.com>   |
| Project B | Company B   | Smith, John /AA <john.smith@companyb.com> |
| Project B | Company E   | Rossi, Luca /AA <luca.rossi@companye.com> |
| Project C | Company F   | Chen, Li /AA <li.chen@companyf.com>       |
| Project C | Company H   | Novak, Eva /AA <eva.novak@companyh.com>   |
| Project C | Company I   | Patel, Riya /AA <riya.patel@companyi.com> |
| Project C | Company A   | Nagy, Anna /AA <anna.nagy@companya.com>   |
| Project C | Company O   | Müller, Lukas /AA <lukas.mueller@companyo.com> |
| Project C | Company P   | Ahmed, Sara /AA <sara.ahmed@companyp.com> |
| Project C | Company Q   | Johnson, Emily /AA <emily.johnson@companyq.com> |
| Project C | Company S   | Dubois, Claire /AA <claire.dubois@companys.com> |
| Project C | Company T   | Tanaka, Hiro /AA <hiro.tanaka@companyt.com> |
| Project D | Company A   | Nagy, Anna /AA <anna.nagy@companya.com>   |


*The contact available contained the information text in one cell, combining the Last Name, First Name & E-mail address, based on the pattern shown in the Contact column.*

In order to generate the Last Name, First Name and E-mail address columns for our mail merge data, 3 excel functions were used to extract each portion of text from the *Contact* column.

**Last Name**
`=TRIM(TEXTBEFORE(E2, ","))`

**First Name**
`=LET(
  s, E2,
  afterComma, TEXTAFTER(s, "","", 1),
  cutAt, IFERROR(TEXTBEFORE(afterComma, "" /""), TEXTBEFORE(afterComma, ""<"")),
  TRIM(cutAt)
)`

**E-mail**
`=TEXTBEFORE(TEXTAFTER(E2, "<"), ">")`



### Output


| Project   | Stakeholder | Last Name | First Name | E-mail                          |
|-----------|-------------|-----------|------------|--------------------------------|
| Project A | Company A   | Nagy      | Anna       | anna.nagy@companya.com         |
| Project A | Company B   | Smith     | John       | john.smith@companyb.com        |
| Project A | Company C   | Kovacs    | Peter      | peter.kovacs@companyc.com      |
| Project A | Company D   | Garcia    | Maria      | maria.garcia@companyd.com      |
| Project A | Company E   | Rossi     | Luca       | luca.rossi@companye.com        |
| Project B | Company A   | Nagy      | Anna       | anna.nagy@companya.com         |
| Project B | Company B   | Smith     | John       | john.smith@companyb.com        |
| Project B | Company E   | Rossi     | Luca       | luca.rossi@companye.com        |
| Project C | Company F   | Chen      | Li         | li.chen@companyf.com           |
| Project C | Company H   | Novak     | Eva        | eva.novak@companyh.com         |
| Project C | Company I   | Patel     | Riya       | riya.patel@companyi.com        |
| Project C | Company A   | Nagy      | Anna       | anna.nagy@companya.com         |
| Project C | Company O   | Müller    | Lukas      | lukas.mueller@companyo.com     |
| Project C | Company P   | Ahmed     | Sara       | sara.ahmed@companyp.com        |
| Project C | Company Q   | Johnson   | Emily      | emily.johnson@companyq.com     |
| Project C | Company S   | Dubois    | Claire     | claire.dubois@companys.com     |
| Project C | Company T   | Tanaka    | Hiro       | hiro.tanaka@companyt.com       |
| Project D | Company A   | Nagy      | Anna       | anna.nagy@companya.com         |





## Resources
The following resources were of chief importance in assembling this project:
- [How to Mail Merge in Microsoft Word](https://www.wikihow.com/Mail-Merge-in-Microsoft-Word)
- [Word: Mail merge by LearnFree](https://www.youtube.com/watch?v=do9ujnZLIC4&embeds_referring_euri=https%3A%2F%2Fwww.bing.com%2F&embeds_referring_origin=https%3A%2F%2Fwww.bing.com&source_ve_path=MjM4NTE)


