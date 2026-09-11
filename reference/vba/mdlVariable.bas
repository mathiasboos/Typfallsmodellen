Attribute VB_Name = "mdlVariable"


'*****************Variables utanför Go_VBA which are used in all modules
Public rngTranslateRange As Range

Public pblnCloseOrSave As Boolean
Public lngTotTasks As Long
Public PctDone As Double
Public lngDoneTask As Long

Public blnOpeningProgress As Boolean
Public nypropp As Integer

Public lngMarginal As Long
Public lngBorn As Double
Public lngWstart As Double
Public lngWTime As Long
Public lngWRef As Long
Public lngDefenitivt_Ar As Long
Public blnNominal As Boolean
'Public par As Double
Public dblPar As Double

Public lngAgeMax As Integer
Public Max_Par_65 As Double
Public dbl_Index_t As Double
Public dbl_Index_t_2 As Double

Public dblBegravningsavgift As Double
Public dblKommunalskatt As Double
Public dblInkomstindex_kvot As Double
Public intFeerules_Year As Integer
Public intMarried As Integer
Public Inkomstindex_uppräkning As Integer
Public dblMakens_Inkomst As Double
Public dblFormogenhet As Double
Public intBoundray_Year As Integer

Public int_Tjanste_PA_KL_Ar As Integer

Public intKvoten As Integer

Public intVal_TJP As Integer
Public IPS_year As Integer
Public intF_enklad_berakning As Integer

Public uttagIP As Double

Public atp_points As Double
Public ATP_earn_years As Long

Public Garanti_ATP_points As Double
Public Garanti_ATP_earn_years As Long
Public Garanti_ATP_min_Ar As Long

Public dblAndel As Double
Public Brutto_Pens As Double
Public GoPens As Long
Public Pens_yearafter As Long

'Public dblGarp As Double
Public dblTP As Double
Public intTP_Faktor As Integer
Public intTP_Faktor_max As Integer
Public dbl_PBB As Double
Public dbl_IBB As Double
Public int_Forsäkringstid_vid_65 As Integer
Public UttagIP_index As Double
Public intPens_Month As Integer
Public lngMonth_index As Long

Public int_Index_Akassa As Integer
Public int_Index_Sjukpenning As Integer
Public int_Index_Sjuklon As Integer
Public int_Index_Foraldrarpenning As Integer

Public Marginal_Reseavdrag As Integer
Public Marginal_Reseavdrag_2 As Integer
Public lngPar_partiellt As Double

Public lngBarnAge_1 As Long
Public lngBarnAge_2 As Long
Public lngBarnAge_3 As Long
Public lngBarnAge_4 As Long

Public dblBarn_inkomst_1 As Double
Public dblBarn_inkomst_2 As Double
Public dblBarn_inkomst_3 As Double
Public dblBarn_inkomst_4 As Double
Public lngKapital_pens As Long
Public intBaldMake As Integer

Public intForvarvsvillkoret As Integer
'Public BruttoI122 As Double
'Public BruttoG122 As Double
'Public BruttoH122 As Double
'Public BruttoBG122 As Double


Public dbl_Yearly_Inflation_Default As Double
Public dbl_Real_Growth_Default As Double
Public dbl_FondAvkastning_Default As Double
Public dbl_Forskottsranta As Double
Public dbl_FondAvkastning As Double

Public lngOffsetRow_NyckelTal As Integer
Public lngOffsetRow_Nagratal As Integer
Public lng_Arvsvinster_TJP As Integer

Public Arr()
Public Arr_Tjanste()
Public Arr_NyckelTal()
Public Arr_NagraTal()
Public Arr_BiIndex()

Public lngOffsetRow As Integer
Public blnCaculateInProcess As Boolean
Public blnGetTyfall As Boolean
Public blnRemoveTypfall As Boolean
Public cMonth_pens As Double
Public ap As Integer
Public apm As Integer

