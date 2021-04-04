VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPermiso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccion de Permisos por Departamentos"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8385
   Begin MSDataListLib.DataList List1 
      Bindings        =   "FrmPermisoxDepartamento.frx":0000
      Height          =   1035
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1826
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Departamento"
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Bindings        =   "FrmPermisoxDepartamento.frx":001E
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8705
      _LayoutType     =   4
      _RowHeight      =   19
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Fecha Desde"
      Columns(0).DataField=   "FechaDesde"
      Columns(0).NumberFormat=   "Edit Mask"
      Columns(0).EditMask=   "##/##/####"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Fecha Hasta"
      Columns(1).DataField=   "FechaHasta"
      Columns(1).NumberFormat=   "Edit Mask"
      Columns(1).EditMask=   "##/##/####"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Departamento"
      Columns(2).DataField=   "Departamento"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   80
      Columns(3)._MaxComboItems=   5
      Columns(3).ValueItems(0)._DefaultItem=   0
      Columns(3).ValueItems(0).Value=   "0"
      Columns(3).ValueItems(0).Value.vt=   8
      Columns(3).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(3).ValueItems(0).DisplayValue(0)=   "bHQAAGoIAABCTWoIAAAAAAAANgAAACgAAAAcAAAAGQAAAAEAGAAAAAAANAgAAAAAAAAAAAAAAAAA"
      Columns(3).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(8)=   "//////////////////////////////////////////////////////////////////+EhoSEhoT/"
      Columns(3).ValueItems(0).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(10)=   "//////////////////////8AAP8AAIQAAISEhoT///////////////////8AAP+EhoT/////////"
      Columns(3).ValueItems(0).DisplayValue(11)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAISE"
      Columns(3).ValueItems(0).DisplayValue(12)=   "hoT///////////8AAP8AAIQAAISEhoT/////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(13)=   "//////////////////8AAP8AAIQAAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAIQAAISEhoT/////"
      Columns(3).ValueItems(0).DisplayValue(14)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAIQA"
      Columns(3).ValueItems(0).DisplayValue(15)=   "AISEhoQAAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(16)=   "//////////////////////8AAP8AAIQAAIQAAIQAAIQAAIQAAIQAAIQAAISEhoT/////////////"
      Columns(3).ValueItems(0).DisplayValue(17)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
      Columns(3).ValueItems(0).DisplayValue(18)=   "AIQAAIQAAIQAAISEhoT/////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(19)=   "//////////////////////////8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(20)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
      Columns(3).ValueItems(0).DisplayValue(21)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(22)=   "//////////////////8AAP8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(23)=   "//////////////////////////////////////////////////8AAP8AAIQAAIQAAISEhoQAAIQA"
      Columns(3).ValueItems(0).DisplayValue(24)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(25)=   "//////8AAP8AAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAISEhoT/////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(26)=   "//////////////////////////////////////////8AAP8AAIQAAISEhoT///////////8AAP8A"
      Columns(3).ValueItems(0).DisplayValue(27)=   "AIQAAIQAAISEhoT/////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(28)=   "//////8AAP8AAIT///////////////////8AAP8AAIQAAIQAAIT/////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(29)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(3).ValueItems(0).DisplayValue(30)=   "AP8AAIQAAP//////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(31)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(36)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(0).DisplayValue(37)=   "//////////////////////////////////////////////////////////////////////8="
      Columns(3).ValueItems(0).DisplayValue.vt=   9
      Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems(1)._DefaultItem=   0
      Columns(3).ValueItems(1).Value=   "1"
      Columns(3).ValueItems(1).Value.vt=   8
      Columns(3).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(3).ValueItems(1).DisplayValue(0)=   "bHQAABYIAABCTRYIAAAAAAAANgAAACgAAAAcAAAAGAAAAAEAGAAAAAAA4AcAAAAAAAAAAAAAAAAA"
      Columns(3).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(10)=   "//////////////////////////////////////+EAACEAAD/////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(11)=   "//////////////////////////////////////////////////////////////////////+EAAAA"
      Columns(3).ValueItems(1).DisplayValue(12)=   "hgAAhgCEAAD/////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(13)=   "//////////////////////////+EAAAAhgAAhgAAhgAAhgCEAAD/////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(14)=   "//////////////////////////////////////////////////////////+EAAAAhgAAhgAAhgAA"
      Columns(3).ValueItems(1).DisplayValue(15)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(16)=   "//////////////+EAAAAhgAAhgAAhgAA/wAAhgAAhgAAhgAAhgCEAAD/////////////////////"
      Columns(3).ValueItems(1).DisplayValue(17)=   "//////////////////////////////////////////////////8AhgAAhgAAhgAA/wD///8A/wAA"
      Columns(3).ValueItems(1).DisplayValue(18)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(19)=   "//////////8A/wAAhgAA/wD///////////8A/wAAhgAAhgAAhgCEAAD/////////////////////"
      Columns(3).ValueItems(1).DisplayValue(20)=   "//////////////////////////////////////////////////8A/wD///////////////////8A"
      Columns(3).ValueItems(1).DisplayValue(21)=   "/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(22)=   "//////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////////"
      Columns(3).ValueItems(1).DisplayValue(23)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(24)=   "//8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(25)=   "//////////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////"
      Columns(3).ValueItems(1).DisplayValue(26)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(27)=   "//////8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(28)=   "//////////////////////////////////////////////8A/wAAhgAAhgCEAAD/////////////"
      Columns(3).ValueItems(1).DisplayValue(29)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(30)=   "//////////8A/wAAhgAAhgD/////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(31)=   "//////////////////////////////////////////////////8A/wD/////////////////////"
      Columns(3).ValueItems(1).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(3).ValueItems(1).DisplayValue(36)=   "//////////////////////////////////8="
      Columns(3).ValueItems(1).DisplayValue.vt=   9
      Columns(3).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems.Count=   2
      Columns(3).Caption=   "Exceso"
      Columns(3).DataField=   "Exceso"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   80
      Columns(4)._MaxComboItems=   5
      Columns(4).ValueItems(0)._DefaultItem=   0
      Columns(4).ValueItems(0).Value=   "0"
      Columns(4).ValueItems(0).Value.vt=   8
      Columns(4).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(4).ValueItems(0).DisplayValue(0)=   "bHQAAGoIAABCTWoIAAAAAAAANgAAACgAAAAcAAAAGQAAAAEAGAAAAAAANAgAAAAAAAAAAAAAAAAA"
      Columns(4).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(8)=   "//////////////////////////////////////////////////////////////////+EhoSEhoT/"
      Columns(4).ValueItems(0).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(10)=   "//////////////////////8AAP8AAIQAAISEhoT///////////////////8AAP+EhoT/////////"
      Columns(4).ValueItems(0).DisplayValue(11)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAISE"
      Columns(4).ValueItems(0).DisplayValue(12)=   "hoT///////////8AAP8AAIQAAISEhoT/////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(13)=   "//////////////////8AAP8AAIQAAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAIQAAISEhoT/////"
      Columns(4).ValueItems(0).DisplayValue(14)=   "//////////////////////////////////////////////////////////8AAP8AAIQAAIQAAIQA"
      Columns(4).ValueItems(0).DisplayValue(15)=   "AISEhoQAAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(16)=   "//////////////////////8AAP8AAIQAAIQAAIQAAIQAAIQAAIQAAIQAAISEhoT/////////////"
      Columns(4).ValueItems(0).DisplayValue(17)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
      Columns(4).ValueItems(0).DisplayValue(18)=   "AIQAAIQAAIQAAISEhoT/////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(19)=   "//////////////////////////8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(20)=   "//////////////////////////////////////////////////////////////8AAP8AAIQAAIQA"
      Columns(4).ValueItems(0).DisplayValue(21)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(22)=   "//////////////////8AAP8AAIQAAIQAAIQAAIQAAISEhoT/////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(23)=   "//////////////////////////////////////////////////8AAP8AAIQAAIQAAISEhoQAAIQA"
      Columns(4).ValueItems(0).DisplayValue(24)=   "AIQAAISEhoT/////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(25)=   "//////8AAP8AAIQAAIQAAISEhoT///8AAP8AAIQAAIQAAISEhoT/////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(26)=   "//////////////////////////////////////////8AAP8AAIQAAISEhoT///////////8AAP8A"
      Columns(4).ValueItems(0).DisplayValue(27)=   "AIQAAIQAAISEhoT/////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(28)=   "//////8AAP8AAIT///////////////////8AAP8AAIQAAIQAAIT/////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(29)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(4).ValueItems(0).DisplayValue(30)=   "AP8AAIQAAP//////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(31)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(36)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(0).DisplayValue(37)=   "//////////////////////////////////////////////////////////////////////8="
      Columns(4).ValueItems(0).DisplayValue.vt=   9
      Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems(1)._DefaultItem=   0
      Columns(4).ValueItems(1).Value=   "1"
      Columns(4).ValueItems(1).Value.vt=   8
      Columns(4).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(4).ValueItems(1).DisplayValue(0)=   "bHQAABYIAABCTRYIAAAAAAAANgAAACgAAAAcAAAAGAAAAAEAGAAAAAAA4AcAAAAAAAAAAAAAAAAA"
      Columns(4).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(2)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(3)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(4)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(5)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(6)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(7)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(8)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(10)=   "//////////////////////////////////////+EAACEAAD/////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(11)=   "//////////////////////////////////////////////////////////////////////+EAAAA"
      Columns(4).ValueItems(1).DisplayValue(12)=   "hgAAhgCEAAD/////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(13)=   "//////////////////////////+EAAAAhgAAhgAAhgAAhgCEAAD/////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(14)=   "//////////////////////////////////////////////////////////+EAAAAhgAAhgAAhgAA"
      Columns(4).ValueItems(1).DisplayValue(15)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(16)=   "//////////////+EAAAAhgAAhgAAhgAA/wAAhgAAhgAAhgAAhgCEAAD/////////////////////"
      Columns(4).ValueItems(1).DisplayValue(17)=   "//////////////////////////////////////////////////8AhgAAhgAAhgAA/wD///8A/wAA"
      Columns(4).ValueItems(1).DisplayValue(18)=   "hgAAhgAAhgCEAAD/////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(19)=   "//////////8A/wAAhgAA/wD///////////8A/wAAhgAAhgAAhgCEAAD/////////////////////"
      Columns(4).ValueItems(1).DisplayValue(20)=   "//////////////////////////////////////////////////8A/wD///////////////////8A"
      Columns(4).ValueItems(1).DisplayValue(21)=   "/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(22)=   "//////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////////"
      Columns(4).ValueItems(1).DisplayValue(23)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(24)=   "//8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(25)=   "//////////////////////////////////////////8A/wAAhgAAhgAAhgCEAAD/////////////"
      Columns(4).ValueItems(1).DisplayValue(26)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(27)=   "//////8A/wAAhgAAhgAAhgCEAAD/////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(28)=   "//////////////////////////////////////////////8A/wAAhgAAhgCEAAD/////////////"
      Columns(4).ValueItems(1).DisplayValue(29)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(30)=   "//////////8A/wAAhgAAhgD/////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(31)=   "//////////////////////////////////////////////////8A/wD/////////////////////"
      Columns(4).ValueItems(1).DisplayValue(32)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(33)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(34)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(35)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(4).ValueItems(1).DisplayValue(36)=   "//////////////////////////////////8="
      Columns(4).ValueItems(1).DisplayValue.vt=   9
      Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems.Count=   2
      Columns(4).Caption=   "Hora Extra"
      Columns(4).DataField=   "Hext"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "CodDepartamento"
      Columns(5).DataField=   "CodDepartamento"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=36,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Named:id=33:Normal"
      _StyleDefs(61)  =   ":id=33,.parent=0"
      _StyleDefs(62)  =   "Named:id=34:Heading"
      _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=34,.wraptext=-1"
      _StyleDefs(65)  =   "Named:id=35:Footing"
      _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   "Named:id=36:Selected"
      _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=37:Caption"
      _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(71)  =   "Named:id=38:HighlightRow"
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
   Begin SmartButtonProject.SmartButton SmartButton1 
      Height          =   855
      Left            =   7200
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      Caption         =   "Salir"
      Picture         =   "FrmPermisoxDepartamento.frx":0037
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoBusca 
      Height          =   375
      Left            =   360
      Top             =   7680
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoBusca"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoDepartamento 
      Height          =   375
      Left            =   4920
      Top             =   6960
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoDepartamento"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoPermiso 
      Height          =   375
      Left            =   360
      Top             =   6960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoPermiso"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmPermiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
With Me.adoPermiso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT dbo.PermisoCalculo.FechaDesde,dbo.PermisoCalculo.FechaHasta, dbo.departamento.departamento,dbo.PermisoCalculo.Exceso, dbo.PermisoCalculo.Hext ,  dbo.PermisoCalculo.CodDepartamento FROM dbo.Departamento INNER JOIN dbo.PermisoCalculo ON dbo.Departamento.CodDepartamento = dbo.PermisoCalculo.CodDepartamento"
   .Refresh
End With

With Me.AdoDepartamento
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT TOP 100 PERCENT CodDepartamento, Departamento From dbo.departamento ORDER BY Departamento"
   .Refresh
End With

With Me.AdoBusca
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

Me.TDBGrid1.Columns(2).Button = True

End Sub

Private Sub List1_DblClick()
Me.TDBGrid1.Columns(2).Text = Me.List1.Text
Me.List1.Visible = False
End Sub

Private Sub List1_LostFocus()
Me.TDBGrid1.Columns(2).Text = Me.List1.Text
Me.List1.Visible = False
End Sub

Private Sub SmartButton1_Click()
Unload Me
End Sub

Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
     Criterio = Me.TDBGrid1.Columns(2)
      Me.AdoBusca.RecordSource = "SELECT TOP 100 PERCENT CodDepartamento, Departamento From dbo.departamento WHERE (Departamento = '" & Criterio & "')ORDER BY Departamento"
      Me.AdoBusca.Refresh
      If Not Me.AdoBusca.Recordset.EOF Then
       Me.TDBGrid1.Columns(5).Text = Me.AdoBusca.Recordset("CodDepartamento")
      End If
End Sub

Private Sub TDBGrid1_BeforeUpdate(Cancel As Integer)
      Criterio = Me.TDBGrid1.Columns(2)
      Me.AdoBusca.RecordSource = "SELECT TOP 100 PERCENT CodDepartamento, Departamento From dbo.departamento WHERE (Departamento = '" & Criterio & "')ORDER BY Departamento"
      Me.AdoBusca.Refresh
      If Not Me.AdoBusca.Recordset.EOF Then
       Me.TDBGrid1.Columns(5).Text = Me.AdoBusca.Recordset("CodDepartamento")
      End If
 
End Sub

Private Sub TDBGrid1_ButtonClick(ByVal ColIndex As Integer)

Select Case ColIndex
 Case 2
 
      Criterio = Me.TDBGrid1.Columns(ColIndex)
      Me.AdoBusca.RecordSource = "SELECT TOP 100 PERCENT CodDepartamento, Departamento From dbo.departamento WHERE (Departamento = '" & Criterio & "')ORDER BY Departamento"
      Me.AdoBusca.Refresh
      If Not Me.AdoBusca.Recordset.EOF Then
       Descripcion = Me.AdoBusca.Recordset("Departamento")
      End If
      Set c = Me.TDBGrid1.Columns(ColIndex)
      With List1
      .Left = Me.TDBGrid1.Left + c.Left
      .Top = Me.TDBGrid1.Top + Me.TDBGrid1.RowTop(Me.TDBGrid1.Row) + Me.TDBGrid1.RowHeight
'      .Width = c.Width + 15
      .Visible = True
      .SetFocus
      .BoundText = Descripcion
      End With
      ColIndexs = 4
  End Select
End Sub

