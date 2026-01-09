# Mail_Merge_using_Excel_and_Word
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


