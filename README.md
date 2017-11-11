# utl_convert_a_sas_dataset_to_excel_without_sas_or_excel_only_need_R
Convert a SAS dataset to excel without SAS or excel (only need R)

    ```  Convert a SAS dataset to excel without SAS or excel (only need R)                                                                                            ```
    ```                                                                                                                                                               ```
    ```  Python and Perl can also read sas7bdat and write to excel.                                                                                                   ```
    ```                                                                                                                                                               ```
    ```  inspired by                                                                                                                                                  ```
    ```  https://goo.gl/HuxBtO                                                                                                                                        ```
    ```  https://communities.sas.com/t5/Base-SAS-Programming/How-to-import-Excel-data-without-ACCESS-to-PC-Files/m-p/325568                                           ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  HAVE SAS DATASET                                                                                                                                             ```
    ```  ================                                                                                                                                             ```
    ```                                                                                                                                                               ```
    ```  Up to 40 obs from sd1.class total obs=6                                                                                                                      ```
    ```                                                                                                                                                               ```
    ```  Obs    NAME                 AGE                                                                                                                              ```
    ```                                                                                                                                                               ```
    ```   1     Alfred                14                                                                                                                              ```
    ```   2     Alice                 13                                                                                                                              ```
    ```   3     Barbara               13                                                                                                                              ```
    ```   4     Carol                 14                                                                                                                              ```
    ```   5     Henry                 14                                                                                                                              ```
    ```   6     James                 12                                                                                                                              ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  WANT (EXCEL SHEET)                                                                                                                                           ```
    ```                                                                                                                                                               ```
    ```  Sheet HAVE                                                                                                                                                   ```
    ```                                                                                                                                                               ```
    ```   +-------------------------+                                                                                                                                 ```
    ```   |      |    A      |   B  |                                                                                                                                 ```
    ```   +------+------------------+                                                                                                                                 ```
    ```   |      |           |      |                                                                                                                                 ```
    ```   |    1 |   NAME    |   AGE|                                                                                                                                 ```
    ```   |    2 |   Alfred  |   14 |                                                                                                                                 ```
    ```   |    3 |   Alice   |   13 |                                                                                                                                 ```
    ```   |    4 |   Barbara |   13 |                                                                                                                                 ```
    ```   |    5 |   Carol   |   14 |                                                                                                                                 ```
    ```   |    6 |   Henry   |   14 |                                                                                                                                 ```
    ```   +------------------+------+                                                                                                                                 ```
    ```                                                                                                                                                               ```
    ```  SOLUTION WORKING CODE ( CAN RECODE IN IML INTERFACE TO R)                                                                                                    ```
    ```                                                                                                                                                               ```
    ```    R  writeWorksheet(wb,have,sheet="have");                                                                                                                   ```
    ```                                                                                                                                                               ```
    ```  FULL SOLUTION                                                                                                                                                ```
    ```  ==============                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  * CREATE SOME DATA;                                                                                                                                          ```
    ```  options validvarname=upcase;                                                                                                                                 ```
    ```  libname sd1 "d:/sd1";                                                                                                                                        ```
    ```  data sd1.class;                                                                                                                                              ```
    ```    set sashelp.class(keep=name age obs=6);                                                                                                                    ```
    ```  run;quit;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```  %utl_submit_r64('                                                                                                                                            ```
    ```  library(haven);                                                                                                                                              ```
    ```  library(XLConnect);                                                                                                                                          ```
    ```  have<-read_sas("d:/sd1/class.sas7bdat");                                                                                                                     ```
    ```  wb <- loadWorkbook("d:/xls/classout.xlsx",create = TRUE);                                                                                                    ```
    ```  createSheet(wb, name="have");                                                                                                                                ```
    ```  writeWorksheet(wb,have,sheet="have");                                                                                                                        ```
    ```  saveWorkbook(wb);                                                                                                                                            ```
    ```  run;quit;                                                                                                                                                    ```
    ```  ');                                                                                                                                                          ```
    ```                                                                                                                                                               ```

