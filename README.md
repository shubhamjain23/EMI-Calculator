# EMI-Calculator

#### An EMI Calculator runs on Excel VBA Macro that calculates and compares EMIs of loans from various banks

It contains two files named `Bank Details.xlsx` and `EMI Calculator.xlsm`

`Bank Details.xlsx` is a source file in which various bank names are stored with their respective interest charges (%) and processing fees.

`EMI Calculator.xlsm` is the main file where you can access the **EMI Calculator** by following the instructions given in that file *or* by just pressing `Ctrl-q` in that workbook.

This macro has 2 UserForms, 1 ClassFile and 14 Modules.

## UserForms
1. **calc_ufm** : 
![EMI Calculator] (https://github.com/shubhamjain23/EMI-Calculator/blob/master/VBAProject%20(EMI%20Calculator.xlsm)/Forms/Mail_Ufm.PNG)

2. **Mail_Ufm** :

## ClassFiles
1. **ThisWorkbook.cls** : It is just used to show a prompt about enabling a set of macro settings whenever user open this workbook.

## Modules
1. **References_Mod.bas** : It is used to check if proper references are loaded in the workbook. (e.g. Mailing options won't work if *Microsoft Outlook* reference is not added.

2. **OpenCalculator_Mod.bas** : It is used for showing the calculator form to user.
3. **UpdateBanks_Mod.bas** : It is used to source data from `Bank Details.xlsx` and save it in `Bank Details` worksheet of `EMI Calculator.xlsm` workbook. 
4. **UpdateBanksLbx_Mod.bas** : It contains a function that populate the *list box* with the banks from `Bank Details` worksheet.

5. **CalcEMI_Mod.bas** : It calculates EMI of all the selected banks and other inputs given by user and make a separate worksheet `Selected Banks` for them.
6. **SortEMI_Mod.bas** : This module sorts EMIs of *Selected Banks* by calling *QuickSort* function.
7. **CreateGraph_Mod.bas** : It is used to create *Bar Chart* of all the banks from `Selected Banks` worksheet and save it in `EMI Graphs` worksheet.

8. **Print_Mod.bas** : It opens the *Print Preview* page for the report generated by user.
9. **SaveReport_Mod.bas** : It is used to save the user generated report in *Report_ddmmyyyy-hhmm.pdf* format.
10. **Email_Mod.bas** : It enables the *Mail Options* to the user, by which user can easily mail report to anyone by specifying *Email-To, Subject and Body*.

11. **QuickSort_Mod.bas** : A module that sorts EMIs of *Selected Banks* bu using QuickSort Algorithm.
12. **PatternMatching_Mod.bas** : It is used for pattern matching used in *Email* and *other variables* by using RegEx.
13. **Swap_Mod.bas** : It contains a function to swap the values of two variables
14. **createSheet_Mod.bas** : It is used to create a Excel Worksheet if it is not already present.
15. **fitAndFormat_Mod.bas** : It contains a function that auto-fit all cells in the worksheet and change the formatting of the first row as specified.
