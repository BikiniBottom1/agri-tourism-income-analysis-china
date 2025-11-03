********************************************************************************
* 农文旅参与对农户收入影响 - 内生转换模型 (ESR) 
* 作者: Qian_Zhou
* 日期: 2025-10-31
********************************************************************************

clear all
set more off
set linesize 120

* 设置工作路径
cd "D:\data"

* 导入数据
import excel "structured_data.xlsx", sheet("Sheet1") firstrow clear

********************************************************************************
* 第一部分：数据准备与描述性统计
********************************************************************************

* 生成因变量：家庭人均纯收入对数
gen lnincome_pc = ln(income/f_size)
label variable lnincome_pc "家庭人均纯收入(对数)"

* 工具变量
gen iv_training = training
label variable iv_training "是否接受技能培训"
gen iv_policy = policy
label variable iv_policy "政策扶持力度"

* 描述性统计（按参与状态）
tabstat lnincome_pc income f_size age_cat edu l_size migrant land_cat, ///
    by(participate) statistics(mean sd min max) columns(statistics) format(%9.2f)

* 参与率
tab participate, missing
proportion participate

********************************************************************************
* 第二部分：内生转换模型(ESR)估计
********************************************************************************

* ESR模型估计
movestay ///
    (lnincome_pc = gender age_cat edu f_size l_size migrant land_cat) ///  /* 参与者收入方程 */
    (lnincome_pc = gender age_cat edu f_size l_size migrant land_cat) ///  /* 不参与者收入方程 */
    , select(participate = gender age_cat edu f_size l_size migrant land_cat iv_training iv_policy) ///
      vce(robust)

* 保存模型结果
estimates store esr_model

* 输出主要结果表
esttab esr_model using "ESR_results.csv", replace ///
    b(%9.3f) se(%9.3f) star(* 0.1 ** 0.05 *** 0.01) ///
    stats(N ll chi2, fmt(0 2 2) labels("样本量" "对数似然" "Chi2")) ///
    title("ESR模型主要结果") ///
    label
/

********************************************************************************
* 第三部分：计算处理效应 ATT & ATU
********************************************************************************

* 预测不同情景下的收入
predict yhat_1, yc1    // 参与者预测收入
predict yhat_0, yc0    // 不参与者预测收入

* 计算 ATT (参与者平均处理效应)
summ yhat_1 if participate==1
scalar E_Y1_D1 = r(mean)
summ yhat_0 if participate==1
scalar E_Y0_D1 = r(mean)
scalar ATT = E_Y1_D1 - E_Y0_D1
scalar ATT_pct = (exp(ATT)-1)*100
display "ATT (参与者收入增长百分比): " ATT_pct "%"

* 计算 ATU (非参与者潜在处理效应)
summ yhat_1 if participate==0
scalar E_Y1_D0 = r(mean)
summ yhat_0 if participate==0
scalar E_Y0_D0 = r(mean)
scalar ATU = E_Y1_D0 - E_Y0_D0
scalar ATU_pct = (exp(ATU)-1)*100
display "ATU (非参与者潜在收入增长百分比): " ATU_pct "%"

********************************************************************************
* 第四部分：模型诊断
********************************************************************************

* 自选择偏差检验：参与者和非参与者相关系数
display "Rho1 (参与者方程相关系数): " e(rho1)
display "Rho2 (非参与者方程相关系数): " e(rho2)

* 工具变量显著性检验（第一阶段）
probit participate gender age_cat edu f_size l_size migrant land_cat iv_training iv_policy, vce(robust)
test iv_training iv_policy
display "工具变量联合显著性检验 p-value: " r(p)

// ********************************************************************************
// * 第五部分：结果输出
// ********************************************************************************
//
// * 输出回归结果到 Word
// * 需要安装：ssc install outreg2
// outreg2 using "esr_results.doc", replace ///
//     title("内生转换模型估计结果") ///
//     ctitle("ESR模型") ///
//     addstat("ATT (对数差异)", ATT, "ATT (%)", ATT_pct, ///
//             "ATU (对数差异)", ATU, "ATU (%)", ATU_pct) ///
//     label dec(3)
//
// * 输出描述性统计到 Excel
// tabstat lnincome_pc income f_size age_cat edu l_size migrant, ///
//     by(participate) statistics(mean sd) columns(statistics) format(%9.3f) save
// matrix A = r(StatTotal)
// putexcel set "descriptive_stats.xlsx", replace
// putexcel A1 = matrix(A), names
//
// * 输出处理效应总结到 Excel
// putexcel set "treatment_effects.xlsx", replace
// putexcel A1 = "处理效应估计结果"
// putexcel A2 = "ATT (对数差异)"     B2 = ATT
// putexcel A3 = "ATT (百分比变化)"   B3 = ATT_pct
// putexcel A4 = "ATU (对数差异)"     B4 = ATU
// putexcel A5 = "ATU (百分比变化)"   B5 = ATU_pct
//
// display "================== ESR分析完成 =================="
// display "请查看输出文件："
// display "1. esr_results.doc - 模型估计结果"
// display "2. descriptive_stats.xlsx - 描述性统计"
// display "3. treatment_effects.xlsx - 处理效应"
