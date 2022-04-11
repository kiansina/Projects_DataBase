# DataBase


#################################################################### 06.Modify_Database (start)
06.Modify_Database(base).py=>

Organize the main database of Ariston and transform it to the structure we need.
*** BEFORE USING THE CODE replace '\ ' with '\' in 'Items' column.
*** PLEASE KEEP A STANDARD CODIFICATION FOR EXAMPLE WE SHOULD EVADE HAVING BOTH "UK" and "ENGLAND"
*** ABOVE WARNING IS ALSO APPLICABLE FOR THE NAMES. WE SHOULD EVADE HAVING BOTH "Elco Heating Solutions Ltd" and "Elco Heating Solutions Limited"

06.Modify_Database(updates).py=>

Organize and clear the new update files sent by company. It is a preparing process to make the file ready to update the main database.
*** BEFORE USING THE CODE delet enter in columns' heads referred to as "Country (risk location)" , "Buildings (reconstruction value)", "Contents (replacement value)", "Stock (peak value per year)"
*** BEFORE USING THE CODE insert path/date/Currency/output manually inside code

06.Modify_Database(updating base).py=> 

Gives the final updated file with indicators of "V" and "Inv" to show that either the data for that row is updated or not
*** BEFORE USING THE CODE insert path

Examples:

Input excel files:
1) "Property Data Base Ariston 07.02.22" as the main database
2) "Ariston_Property_Location_2022 - Completed" as the new update file

output excel files:
1) "Results" the suitable transformation of Input 1
2) "ds" as the suitable transformation of Input 2
*** Two external lines are added to "ds" manually to show the final results with "V" and "Inv" indicators.
3) dfff The final results.
#################################################################### 06.Modify_Database (end)
