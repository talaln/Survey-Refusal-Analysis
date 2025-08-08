/*=============================================================
 Project : Survey Refusal Analysis (Synthetic)
 Author  : Mohammad Talal Naseem
 Purpose : Generate Tables 1â€“3 using synthetic data matching required variables
 Inputs  : ./data/Refusal_Data.xlsx (sheets: main_dummy, main _cat, scrnr_dummy, scrn_cat)
 Outputs : ./out/refusal_analysis_tables.xlsx and .rtf
==============================================================*/

/*------------------ Parameters ------------------------------*/
%let in_xlsx  = ./data/Refusal_Data.xlsx;
%let out_dir  = ./out;
%let out_xlsx = &out_dir./refusal_analysis_tables.xlsx;
%let out_rtf  = &out_dir./refusal_analysis_tables.rtf;

/*------------------ Setup ----------------------------------*/
options mprint mlogic symbolgen nodate nonumber;
ods _all_ close;
filename reffile "&in_xlsx";

/*------------------ Import ---------------------------------*/
proc import datafile=reffile out=work.main_dummy dbms=xlsx replace;
    sheet="main_dummy"; getnames=yes;
run;

proc import datafile=reffile out=work.main_cat dbms=xlsx replace;
    sheet="main _cat"; getnames=yes; /* note the space */
run;

proc import datafile=reffile out=work.scrnr_dummy dbms=xlsx replace;
    sheet="scrnr_dummy"; getnames=yes;
run;

proc import datafile=reffile out=work.scrnr_cat dbms=xlsx replace;
    sheet="scrn_cat"; getnames=yes;
run;

/*------------------ Formats --------------------------------*/
proc format;
    value resultcodefmt
        1001='Complete'
        1003='Final Refusal'
        1002='Interim Refusal';
run;

/*================== Table 1: Categories (Screener & Main) ===*/
%macro SC(cat, ct);
proc freq data=work.scrnr_cat noprint;
    tables RC*&cat. / nopercent nocum nocol out=scrn_cat_&ct. outpct;
    format RC resultcodefmt.;
run;
%mend;

%SC(General,gen);
%SC(Negative,neg);
%SC(Cultural,cult);
%SC(Survey,surv);

data work.scrn_cat_prev;
    set work.scrn_cat_gen work.scrn_cat_neg work.scrn_cat_cult work.scrn_cat_surv;
run;

data work.scrn_cat_prev(drop=general negative cultural survey percent pct_col);
    length Category $10;
    set work.scrn_cat_prev;
    if RC='' then delete;
    if general=1 then Category='General';
    else if negative=1 then Category='Negative';
    else if cultural=1 then Category='Cultural';
    else if survey=1 then Category='Survey';
    else delete;
run;

%macro MN(cat, ct);
proc freq data=work.main_cat noprint;
    tables RC*&cat. / nopercent nocum nocol out=main_cat_&ct. outpct;
    format RC resultcodefmt.;
run;
%mend;

%MN(General,gen);
%MN(Negative,neg);
%MN(Cultural,cult);
%MN(Survey,surv);
%MN(Health,hth);

data work.main_cat_prev;
    set work.main_cat_gen work.main_cat_neg work.main_cat_cult work.main_cat_surv work.main_cat_hth;
run;

data work.main_cat_prev(drop=general negative cultural survey health percent pct_col);
    length Category $10;
    set work.main_cat_prev;
    if RC='' then delete;
    if general=1 then Category='General';
    else if negative=1 then Category='Negative';
    else if cultural=1 then Category='Cultural';
    else if survey=1 then Category='Survey';
    else if health=1 then Category='Health';
    else delete;
run;

proc logistic data=work.scrnr_cat;
    class RC;
    model RC (event='Interim Refusal') = negative cultural survey / rsquare;
    format RC resultcodefmt.;
    ods output parameterestimates=work.scrn_cat_mle;
run;

data work.scrn_cat_mle(drop=df WaldChiSq _ESTTYPE_);
    set work.scrn_cat_mle;
    rename Variable=Category;
    if upcase(Variable)='INTERCEPT' then delete;
run;

proc logistic data=work.main_cat;
    class RC;
    model RC (event='Interim Refusal') = negative cultural survey health / rsquare;
    format RC resultcodefmt.;
    ods output parameterestimates=work.main_cat_mle;
run;

data work.main_cat_mle(drop=df WaldChiSq _ESTTYPE_);
    set work.main_cat_mle;
    rename Variable=Category;
    if upcase(Variable)='INTERCEPT' then delete;
run;

proc sort data=work.scrn_cat_mle;  by Category; run;
proc sort data=work.scrn_cat_prev; by Category; run;
data work.scrn_cat_T1; merge work.scrn_cat_prev work.scrn_cat_mle; by Category; Type='Screener'; run;

proc sort data=work.main_cat_mle;  by Category; run;
proc sort data=work.main_cat_prev; by Category; run;
data work.main_cat_T1; merge work.main_cat_prev work.main_cat_mle; by Category; Type='Main'; run;

data work.Table1; set work.scrn_cat_T1 work.main_cat_T1; run;

title "Table 1: Refusal Categories by Type with Logistic Estimates";
proc report data=work.Table1 missing nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Type Category RC,(count pct_row) estimate StdErr;
    define Type / group;
    define Category / group;
    define RC / across;
    define count / "Frequency" style(column)=[just=center cellwidth=1in];
    define Pct_row / "Percent" format=8.2 style(column)=[just=center cellwidth=1in];
    define estimate / group;
    define StdErr / group;
run; title;

/*================== Table 2: Screener Multinomial (scrnr_dummy) ===*/
proc logistic data=work.scrnr_dummy;
    class RC;
    model RC (ref='Complete') = Rural Eastern Northern Southern Western Single_Unit Middle High / link=glogit rsquare;
    format RC resultcodefmt.;
    ods output parameterestimates=work.Table2;
run;

title "Table 2: Screener Multinomial Logistic Parameter Estimates";
proc report data=work.Table2 missing nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Variable Response,(estimate StdErr);
    define Variable / group;
    define Response / across;
    define estimate / group;
    define StdErr / group;
run; title;

/*================== Table 3: Main interview multinomial model ======================*/
proc logistic order=data data=work.main_dummy;
    class sResultCodeId Gender_R Gender_I;
    model sResultCodeId (ref='Complete') =
        Rural Eastern Northern Southern Western Single_Unit Middle High 
        '15-18'n '19-34'n '35-49'n 'I_16-18'n 'I_19-34'n 'I_35-49'n
        Gender_R Gender_I Gender_R*Gender_I
        / link=glogit rsquare;
    format sResultCodeId resultcodefmt.;
    ods output parameterestimates=work.Table3_raw;
run;

/* Harmonize columns so PROC REPORT has VARIABLE, RESPONSE, ESTIMATE, STDERR */
data work.Table3;
    length Variable $64;
    set work.Table3_raw;
    /* Some SAS versions output Effect/Parameter instead of Variable */
    if missing(Variable) then Variable = coalescec(Effect, Parameter);
    keep Variable Response Estimate StdErr;
run;

/* PROC REPORT with DISPLAY usage for numeric columns */
title "Table 3: Main Interview Multinomial Logistic Parameter Estimates";
proc report data=work.Table3 missing nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Variable Response Estimate StdErr;
    define Variable / "Predictor" group;
    define Response / "Response Category" group;
    define Estimate / display "Estimate" style(column)=[just=center cellwidth=1in];
    define StdErr  / display "Std. Error" style(column)=[just=center cellwidth=1in];
run;
title;

/*------------------ Outputs --------------------------------*/
ods excel file="&out_xlsx" options(embedded_titles='yes');
title "Table 1: Refusal Categories by Type with Logistic Estimates";
proc report data=work.Table1 missing nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Type Category RC,(count pct_row) estimate StdErr;
    define Type / group;
    define Category / group;
    define RC / across;
    define count / "Frequency" style(column)=[just=center cellwidth=1in];
    define Pct_row / "Percent" format=8.2 style(column)=[just=center cellwidth=1in];
    define estimate / group;
    define StdErr / group;
run;

title "Table 2: Screener Multinomial Logistic Parameter Estimates";
proc report data=work.Table2 missing nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Variable Response,(estimate StdErr);
    define Variable / group;
    define Response / across;
    define estimate / group;
    define StdErr / group;
run;

title "Table 3: Main Interview Multinomial Logistic Parameter Estimates";
proc report data=work.Table3 missing nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Variable Response Estimate StdErr;
    define Variable / "Predictor" group;
    define Response / "Response Category" group;
    define Estimate / display "Estimate" style(column)=[just=center cellwidth=1in];
    define StdErr  / display "Std. Error" style(column)=[just=center cellwidth=1in];
run;
ods excel close; title;

ods rtf file="&out_rtf" startpage=no notoc_data;
title "Table 1: Refusal Categories by Type with Logistic Estimates";
proc report data=work.Table1 missing nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Type Category RC,(count pct_row) estimate StdErr;
    define Type / group;
    define Category / group;
    define RC / across;
    define count / "Frequency" style(column)=[just=center cellwidth=1in];
    define Pct_row / "Percent" format=8.2 style(column)=[just=center cellwidth=1in];
    define estimate / group;
    define StdErr / group;
run;

title "Table 2: Screener Multinomial Logistic Parameter Estimates";
proc report data=work.Table2 missing nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Variable Response,(estimate StdErr);
    define Variable / group;
    define Response / across;
    define estimate / group;
    define StdErr / group;
run;

title "Table 3: Main Interview Multinomial Logistic Parameter Estimates";
proc report data=work.Table3 missing nowd
    style(header)=[bordertopcolor=black borderbottomcolor=black borderrightcolor=white borderleftcolor=white
                   font_face='Times New Roman' font_weight=bold font_size=9pt color=black just=center background=white]
    style(column)=[bordertopcolor=white borderbottomcolor=white borderrightcolor=white borderleftcolor=white
                   just=left font_face='Times New Roman' font_size=8pt];
    column Variable Response Estimate StdErr;
    define Variable / "Predictor" group;
    define Response / "Response Category" group;
    define Estimate / display "Estimate" style(column)=[just=center cellwidth=1in];
    define StdErr  / display "Std. Error" style(column)=[just=center cellwidth=1in];
run;
ods rtf close; title;

