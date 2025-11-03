*------------------------------------------------------------
* 导入数据
*------------------------------------------------------------
import excel "D:\data\structured_data.xlsx", sheet("Sheet1") firstrow clear

* 查看变量基本情况
tab participate
tab age_cat
tab edu
tab land_cat

*------------------------------------------------------------
* 1. 卡方检验（分类变量 vs 参与情况）
*------------------------------------------------------------
tab age_cat participate, chi2
tab edu participate, chi2
tab land_cat participate, chi2

*------------------------------------------------------------
* 3. 若想展示结果在表格中（输出为Excel）
*------------------------------------------------------------
estpost tabulate age_cat participate, chi2
esttab using "D:\data\chi2_results.xlsx", cells("chi2 p") replace
