Attribute VB_Name = "mGUIStrings"
Option Explicit

Private Declare Function GetUserDefaultUILanguage Lib "kernel32.dll" () As Long

Public Enum efnGUIString
    ' Forms
    ' General
    efnGUIStr_General_CloseButton_Caption
    efnGUIStr_General_OKButton_Caption
    efnGUIStr_General_CancelButton_Caption
    efnGUIStr_General_PageNumbersFormatString_Default
    ' frmClipboardCopiedMessage
    efnGUIStr_frmClipboardCopiedMessage_lblMessage_Caption
    ' frmConfigHistory
    efnGUIStr_frmConfigHistory_Caption
    efnGUIStr_frmConfigHistory_chkRememberHistory_Caption
    efnGUIStr_frmConfigHistory_cmdEraseContext_Caption
    efnGUIStr_frmConfigHistory_cmdEraseAll_Caption
    efnGUIStr_frmConfigHistory_HelpMessageTitle
    efnGUIStr_frmConfigHistory_HelpMessage
    ' frmCopyGridTextOptions
    efnGUIStr_frmCopyGridTextOptions_Caption
    efnGUIStr_frmCopyGridTextOptions_lblColumnsSeparationMode_Caption
    efnGUIStr_frmCopyGridTextOptions_cboMode_List
    efnGUIStr_lblSelectComunsToInclude_Caption
    efnGUIStr_EnterColumnSeparatorMessage
    efnGUIStr_EnterColumnSeparatorMessageTitle
    efnGUIStr_SelectFontMessage
    '  frmPrintGridFormatOptions
    efnGUIStr_frmPrintGridFormatOptions_Caption
    efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_0
    efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_1
    efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_2
    efnGUIStr_frmPrintGridFormatOptions_chkEnableAutoOrientation_Caption
    efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersSeparatorLine_Caption
    efnGUIStr_frmPrintGridFormatOptions_chkPrintColumnsHeadersLines_Caption
    efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersBorder_Caption
    efnGUIStr_frmPrintGridFormatOptions_chkPrintFixedColsBackground_Caption
    efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersBackground_Caption
    efnGUIStr_frmPrintGridFormatOptions_chkPrintOtherBackgrounds_Caption
    efnGUIStr_frmPrintGridFormatOptions_chkPrintRowsLines_Caption
    efnGUIStr_frmPrintGridFormatOptions_chkPrintColumnsDataLines_Caption
    efnGUIStr_frmPrintGridFormatOptions_chkPrintOuterBorder_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblLineWidth_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblStyle_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblOtherTextsFont_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblSubheadingFont_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblHeadingFont_Caption
    efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersFont_Caption
    efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersFormat_Caption
    efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersPosition_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblGridAlign_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblColor_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblScalePercent_Caption
    efnGUIStr_frmPrintGridFormatOptions_cboColor_List
    efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_cboPageNumbersPosition_List
    efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_List
    efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_Style
    efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_CustomStyle
    efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_Customize
    efnGUIStr_frmPrintGridFormatOptions_lblSample_Caption
    efnGUIStr_frmPrintGridFormatOptions_DrawSample_Column
    efnGUIStr_frmPrintGridFormatOptions_DrawSample_Element
    efnGUIStr_frmPrintGridFormatOptions_DrawSample_Data
    efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_ToolTipText
    efnGUIStr_frmPrintGridFormatOptions_txtLineWidth_ToolTipText
    efnGUIStr_frmPrintGridFormatOptions_cmdHeadersBackgroundColor_ToolTipText
    efnGUIStr_frmPrintGridFormatOptions_txtLineWidthHeadersSeparatorLine_ToolTipText
    efnGUIStr_frmPrintGridFormatOptions_VariousChangeColorCommandButtons_ToolTipText
    efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidth_Message
    efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidthHeadersSeparatorLine_Message
    ' frmPageNumbersOptions
    efnGUIStr_frmPageNumbersOptions_Caption
    
    ' frmSelectColumns
    efnGUIStr_frmSelectColumns_Caption
    efnGUIStr_frmSelectColumns_lblTitle_Caption
    efnGUIStr_frmSelectColumns_OneVisible_Message
    ' frmSettingGridDataProgress
    efnGUIStr_frmSettingGridDataProgress_lblMessage_Caption_Start
    efnGUIStr_frmSettingGridDataProgress_lblMessage_Caption_Progress
    ' frmPrintPreview
    efnGUIStr_frmPrintPreview_Caption
    efnGUIStr_frmPrintPreview_mnuView2p_Caption
    efnGUIStr_frmPrintPreview_mnuView3p_Caption
    efnGUIStr_frmPrintPreview_mnuView6p_Caption
    efnGUIStr_frmPrintPreview_mnuView12p_Caption
    efnGUIStr_frmPrintPreview_mnuIconsAuto_Caption
    efnGUIStr_frmPrintPreview_mnuIconsSmall_Caption
    efnGUIStr_frmPrintPreview_mnuIconsMedium_Caption
    efnGUIStr_frmPrintPreview_mnuIconsBig_Caption
    efnGUIStr_frmPrintPreview_mnuIconsShowBottomToolBar_Caption
    efnGUIStr_frmPrintPreview_CurrentlySelected
    efnGUIStr_frmPrintPreview_CurrentlyShown
    efnGUIStr_frmPrintPreview_lblPageOrientation_Caption
    efnGUIStr_frmPrintPreview_lblView_Caption
    efnGUIStr_frmPrintPreview_lblScalePercent_Caption
    efnGUIStr_frmPrintPreview_lblPage_Singular_Caption
    efnGUIStr_frmPrintPreview_lblPage_Plural_Caption
    efnGUIStr_frmPrintPreview_lblPageCount_Caption
    efnGUIStr_frmPrintPreview_PreparingDoc_Caption
    efnGUIStr_frmPrintPreview_cmdClose_Caption
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Print
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageSetup
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageNumbersOptions
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Format
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationPortrait
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationLandscape
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewNormalSize
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenWidth
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenHeight
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewSeveralPages
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_DecreaseScale
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_IncreaseScale
    efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_FirstPage
    efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Singular
    efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Plural
    efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Singular
    efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Plural
    efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_LastPage
    
    ' UserControls
    ' FontPicker
    efnGUIStr_FontPicker_ButtonToolTipTextDefault
    efnGUIStr_FontPicker_DrawSample_Bold
    efnGUIStr_FontPicker_DrawSample_Italic
    ' FontSizeChanger
    efnGUIStr_FontSizeChanger_Extender_ToolTipText
    efnGUIStr_FontSizeChanger_mnuDefaultFontSize_Caption
    efnGUIStr_FontSizeChanger_btnMinus_ToolTipText
    efnGUIStr_FontSizeChanger_btnPlus_ToolTipText
    ' FlexFn
    efnGUIStr_FlexFn_PrintButton_ToolTipText_Default
    efnGUIStr_FlexFn_PrintPreviewButton_ToolTipText_Default
    efnGUIStr_FlexFn_FindButton_ToolTipText_Default
    efnGUIStr_FlexFn_CopyButton_ToolTipText_Default
    efnGUIStr_FlexFn_SaveButton_ToolTipText_Default
    efnGUIStr_FlexFn_GroupDataButton_ToolTipText_Default
    efnGUIStr_FlexFn_GroupDataButtonPressed_ToolTipText_Default
    efnGUIStr_FlexFn_ConfigColumnsButton_ToolTipText_Default
    efnGUIStr_FlexFn_ConfigColumnsButtonColsHidden_ToolTipText_Default
    efnGUIStr_FlexFn_CopyCellMenuCaption_Default
    efnGUIStr_FlexFn_CopyRowMenuCaption_Default
    efnGUIStr_FlexFn_CopyColumnMenuCaption_Default
    efnGUIStr_FlexFn_CopyAllMenuCaption_Default
    efnGUIStr_FlexFn_CopySelectionMenuCaption_Default
    efnGUIStr_FlexFn_mnuCopyParent_Caption
    ' History
    efnGUIStr_History_mnuDelete_Caption1
    efnGUIStr_History_mnuDelete_Caption2
    efnGUIStr_History_ToolTipTextStart_Default
    efnGUIStr_History_ToolTipTextSelect_Default
    efnGUIStr_History_BackButtonToolTipText_Default
    efnGUIStr_History_ForwardButtonToolTipText_Default
    ' DateEnter
    efnGUIStr_DateEnter_ToolTipTextStart_Default
    efnGUIStr_DateEnter_ToolTipTextEnd_Default
    efnGUIStr_DateEnter_Validate1_MsgBoxTitle
    efnGUIStr_DateEnter_Validate1_MsgBoxError1
    efnGUIStr_DateEnter_Validate1_MsgBoxError2
    efnGUIStr_DateEnter_Validate1_MsgBoxError3
    efnGUIStr_DateEnter_Validate1_MsgBoxError4
    efnGUIStr_DateEnter_Validate1_MsgBoxError5
    efnGUIStr_DateEnter_Validate1_MsgBoxError6
    efnGUIStr_DateEnter_Validate1_MsgBoxError7
    efnGUIStr_DateEnter_Validate1_MsgBoxError8
    ' Class modules
    ' cGridHandler
    efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString1
    efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString2
    efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString3
    efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString4
    
    ' FlexFnObject
    efnGUIStr_FlexFnObject_FindTextInGrid_InputBoxEnterTextMessageTitle
    efnGUIStr_FlexFnObject_FindTextInGrid_InputBoxEnterTextMessage
    efnGUIStr_FlexFnObject_FindTextInGrid_MsgboxTextNotFound
    efnGUIStr_FlexFnObject_SaveGridAsExcelFile_iDlg_Filter
    efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1a
    efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1b
    efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2a
    efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2b
    efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox3
    efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox4
    efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox5
    efnGUIStr_FlexFnObject_PageNumbersFormatStrings_Count
    efnGUIStr_FlexFnObject_PageNumbersFormatStrings
    ' cPrinterEx
    efnGUIStr_cPrinterEx_PrintDocumentTooLongWarning_MsgBoxWarning
    
    ' Bas modules
    'mGlobals
    efnGUIStr_mGlobals_ValidFileName_DefaultFileName
End Enum

Private mGUILanguage As vbExGUILanguageConstants


Public Function GetLocalizedString(nID As efnGUIString, Optional nIndex As Long, Optional ByVal nLang As vbExGUILanguageConstants) As String
    If mGUILanguage = vxLangAUTO Then SetGUILanguage
    If nLang = vxLangAUTO Then nLang = mGUILanguage
    
    Select Case nLang
        Case vxLangCHINESE_SIMPLIFIED ' Thanks ChenLin: http://www.vbforums.com/showthread.php?865299#post5309543
            Select Case nID
                ' General
                Case efnGUIStr_General_CloseButton_Caption
                    GetLocalizedString = "&C 关闭"
                Case efnGUIStr_General_OKButton_Caption
                    GetLocalizedString = "&O 确定"
                Case efnGUIStr_General_CancelButton_Caption
                    GetLocalizedString = "&C 取消"
                Case efnGUIStr_General_PageNumbersFormatString_Default
                    GetLocalizedString = "#"
                ' frmClipboardCopiedMessage
                Case efnGUIStr_frmClipboardCopiedMessage_lblMessage_Caption
                    GetLocalizedString = "复制文字"
                ' frmConfigHistory
                Case efnGUIStr_frmConfigHistory_Caption
                    GetLocalizedString = "历史配置"
                Case efnGUIStr_frmConfigHistory_chkRememberHistory_Caption
                    GetLocalizedString = "记住历史记录"
                Case efnGUIStr_frmConfigHistory_HelpMessageTitle
                    GetLocalizedString = "历史"
                Case efnGUIStr_frmConfigHistory_HelpMessage
                    GetLocalizedString = "确定程序是否会记住它下一次被搜索或选择的内容，在关闭后您可以通过单击其中一个按钮来删除当前历史，" & _
                                        vbCrLf & vbCrLf & "您也可以通过单击右键来删除历史记录中的一个项目。 鼠标单击（在列表中，在进入此配置屏幕之前）"
                Case efnGUIStr_frmConfigHistory_cmdEraseContext_Caption
                    GetLocalizedString = "删除此上下文的历史记录"
                Case efnGUIStr_frmConfigHistory_cmdEraseAll_Caption
                    GetLocalizedString = "删除全部"
                ' frmCopyGridTextOptions
                Case efnGUIStr_frmCopyGridTextOptions_Caption
                    GetLocalizedString = "复制文字选项"
                Case efnGUIStr_frmCopyGridTextOptions_lblColumnsSeparationMode_Caption
                    GetLocalizedString = "列分隔符:"
                Case efnGUIStr_frmCopyGridTextOptions_cboMode_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "Tab 键（建议用于复制到 Excel 程序使用）"
                        Case 1
                            GetLocalizedString = "相同宽度空格（复制具有固定宽度字体的文本编辑器）"
                        Case 2
                            GetLocalizedString = "根据字体的间距（复制到带有可变宽度字体的文本编辑器）"
                        Case 3
                            GetLocalizedString = "使用自定义字符或文本作为分隔符"
                    End Select
                Case efnGUIStr_lblSelectComunsToInclude_Caption
                    GetLocalizedString = "选择要包含的列："
                Case efnGUIStr_EnterColumnSeparatorMessage
                    GetLocalizedString = "请输入文本或字符作为列分隔符"
                Case efnGUIStr_EnterColumnSeparatorMessageTitle
                    GetLocalizedString = "输入分隔符"
                Case efnGUIStr_SelectFontMessage
                    GetLocalizedString = "请选择将要使用的目标程序使用哪种字体进行粘贴,列的对齐方式为近似值."
                ' frmPrintGridFormatOptions
                Case efnGUIStr_frmPrintGridFormatOptions_Caption
                    GetLocalizedString = "打印格式"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_0
                    GetLocalizedString = "&O 选项"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_1
                    GetLocalizedString = "&S 样式"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_2
                    GetLocalizedString = "&M 更多"
                Case efnGUIStr_frmPrintGridFormatOptions_chkEnableAutoOrientation_Caption
                    GetLocalizedString = "如果报表比纸张宽，则自动将页面方向更改为水平。"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersSeparatorLine_Caption
                    GetLocalizedString = "标题分隔符"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintColumnsHeadersLines_Caption
                    GetLocalizedString = "列标题行"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersBorder_Caption
                    GetLocalizedString = "标题边界色"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintFixedColsBackground_Caption
                    GetLocalizedString = "固定列背景色"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersBackground_Caption
                    GetLocalizedString = "标题背景色"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintOtherBackgrounds_Caption
                    GetLocalizedString = "其他背景色"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintRowsLines_Caption
                    GetLocalizedString = "行线"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintColumnsDataLines_Caption
                    GetLocalizedString = "列线"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintOuterBorder_Caption
                    GetLocalizedString = "外部边缘"
                Case efnGUIStr_frmPrintGridFormatOptions_lblLineWidth_Caption
                    GetLocalizedString = "线宽:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblStyle_Caption
                    GetLocalizedString = "样式:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblOtherTextsFont_Caption
                    GetLocalizedString = "其他文本字体:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblSubheadingFont_Caption
                    GetLocalizedString = "小标题字体:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblHeadingFont_Caption
                    GetLocalizedString = "标题或标题字体："
                Case efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersFont_Caption
                    GetLocalizedString = "页码字体:"
                Case efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersFormat_Caption
                    GetLocalizedString = "页码格式:"
                Case efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersPosition_Caption
                    GetLocalizedString = "页码位置:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblGridAlign_Caption
                    GetLocalizedString = "网格对齐:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblColor_Caption
                    GetLocalizedString = "颜色:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblScalePercent_Caption
                    GetLocalizedString = "样式:"
                Case efnGUIStr_frmPrintGridFormatOptions_cboColor_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "打印机默认"
                        Case 1
                            GetLocalizedString = "灰度"
                        Case 2
                            GetLocalizedString = "彩色"
                    End Select
                Case efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_cboPageNumbersPosition_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "右下角"
                        Case 1
                            GetLocalizedString = "左下角"
                        Case 2
                            GetLocalizedString = "底部居中"
                        Case 3
                            GetLocalizedString = "右上角"
                        Case 4
                            GetLocalizedString = "左上角"
                        Case 5
                            GetLocalizedString = "顶部居中"
                        Case 6
                            GetLocalizedString = "不显示页码"
                    End Select
                Case efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "居中"
                        Case 1
                            GetLocalizedString = "居左"
                        Case 2
                            GetLocalizedString = "居右"
                        Case 3
                            GetLocalizedString = "拉伸"
                    End Select
                Case efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_Style
                    GetLocalizedString = "样式"
                Case efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_CustomStyle
                    GetLocalizedString = "自定义样式"
                Case efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_Customize
                    GetLocalizedString = "定制"
                Case efnGUIStr_frmPrintGridFormatOptions_lblSample_Caption
                    GetLocalizedString = "事例:"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Column
                    GetLocalizedString = "列"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Element
                    GetLocalizedString = "元素"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Data
                    GetLocalizedString = "数据"
                Case efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_ToolTipText
                    GetLocalizedString = "仅在数据网格比页面更窄时才有效"
                Case efnGUIStr_frmPrintGridFormatOptions_txtLineWidth_ToolTipText
                    GetLocalizedString = "改变线条粗细"
                Case efnGUIStr_frmPrintGridFormatOptions_cmdHeadersBackgroundColor_ToolTipText
                    GetLocalizedString = "更改标题（或固定列）的背景颜色"
                Case efnGUIStr_frmPrintGridFormatOptions_txtLineWidthHeadersSeparatorLine_ToolTipText
                    GetLocalizedString = "页眉分隔线高度"
                Case efnGUIStr_frmPrintGridFormatOptions_VariousChangeColorCommandButtons_ToolTipText
                    GetLocalizedString = "修改颜色"
                Case efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidth_Message
                    GetLocalizedString = "线条的宽度值必须在1到80之间"
                Case efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidthHeadersSeparatorLine_Message
                    GetLocalizedString = "线条的宽度值必须在1到80之间"
                ' frmPageNumbersOptions
                Case efnGUIStr_frmPageNumbersOptions_Caption
                    GetLocalizedString = "页码"
                ' frmSelectColumns
                Case efnGUIStr_frmSelectColumns_Caption
                    GetLocalizedString = "配置可见列"
                Case efnGUIStr_frmSelectColumns_lblTitle_Caption
                    GetLocalizedString = "&S 选择要显示的列:"
                Case efnGUIStr_frmSelectColumns_OneVisible_Message
                    GetLocalizedString = "必须至少选择一列。"
                ' frmSettingGridDataProgress
                Case efnGUIStr_frmSettingGridDataProgress_lblMessage_Caption_Start
                    GetLocalizedString = "生成预览"
                Case efnGUIStr_frmSettingGridDataProgress_lblMessage_Caption_Progress
                    GetLocalizedString = "生成预览，总页数"
                ' frmPrintPreview
                Case efnGUIStr_frmPrintPreview_Caption
                    GetLocalizedString = "打印预览"
                Case efnGUIStr_frmPrintPreview_mnuView2p_Caption
                    GetLocalizedString = "显示两页"
                Case efnGUIStr_frmPrintPreview_mnuView3p_Caption
                    GetLocalizedString = "显示三页"
                Case efnGUIStr_frmPrintPreview_mnuView6p_Caption
                    GetLocalizedString = "显示六页"
                Case efnGUIStr_frmPrintPreview_mnuView12p_Caption
                    GetLocalizedString = "显示12页"
                Case efnGUIStr_frmPrintPreview_mnuIconsAuto_Caption
                    GetLocalizedString = "自动"
                Case efnGUIStr_frmPrintPreview_mnuIconsSmall_Caption
                    GetLocalizedString = "小图标"
                Case efnGUIStr_frmPrintPreview_mnuIconsMedium_Caption
                    GetLocalizedString = "中图标"
                Case efnGUIStr_frmPrintPreview_mnuIconsBig_Caption
                    GetLocalizedString = "中等图标"
                Case efnGUIStr_frmPrintPreview_mnuIconsShowBottomToolBar_Caption
                    GetLocalizedString = "显示底部工具栏"
                Case efnGUIStr_frmPrintPreview_CurrentlySelected
                    GetLocalizedString = "已选择" 'selected"
                Case efnGUIStr_frmPrintPreview_CurrentlyShown
                    GetLocalizedString = "正在显示" 'now shown"
                Case efnGUIStr_frmPrintPreview_lblPageOrientation_Caption
                    GetLocalizedString = "显示方向:"
                Case efnGUIStr_frmPrintPreview_lblView_Caption
                    GetLocalizedString = "多页显示:"
                Case efnGUIStr_frmPrintPreview_lblScalePercent_Caption
                    GetLocalizedString = "&S 缩放："
                Case efnGUIStr_frmPrintPreview_lblPage_Singular_Caption
                    GetLocalizedString = "&P 页数:"
                Case efnGUIStr_frmPrintPreview_lblPage_Plural_Caption
                    GetLocalizedString = "&P 页数:" '总页:"
                Case efnGUIStr_frmPrintPreview_lblPageCount_Caption
                    GetLocalizedString = "/"
                Case efnGUIStr_frmPrintPreview_PreparingDoc_Caption
                    GetLocalizedString = "正在生成打印预览..."
                Case efnGUIStr_frmPrintPreview_cmdClose_Caption
                    GetLocalizedString = "&C 关闭"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Print
                    GetLocalizedString = "开始打印"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageSetup
                    GetLocalizedString = "页面设置"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageNumbersOptions
                    GetLocalizedString = "页码"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Format
                    GetLocalizedString = "显示格式"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationPortrait
                    GetLocalizedString = "纵向显示"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationLandscape
                    GetLocalizedString = "横向显示"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewNormalSize
                    GetLocalizedString = "查看正常页面大小"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenWidth
                    GetLocalizedString = "页面调整为屏幕宽度"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenHeight
                    GetLocalizedString = "页面调整到屏幕高度"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewSeveralPages
                    GetLocalizedString = "显示多页"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_DecreaseScale
                    GetLocalizedString = "减少字体和元素大小"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_IncreaseScale
                    GetLocalizedString = "增加字体和元素大小"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_FirstPage
                    GetLocalizedString = "首页"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Singular
                    GetLocalizedString = "上一页"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Plural
                    GetLocalizedString = "上一个多页"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Singular
                    GetLocalizedString = "下一页"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Plural
                    GetLocalizedString = "下一个多页"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_LastPage
                    GetLocalizedString = "最后页"
                ' UserControls
                ' FontPicker
                Case efnGUIStr_FontPicker_ButtonToolTipTextDefault
                    GetLocalizedString = "选择字体"
                Case efnGUIStr_FontPicker_DrawSample_Bold
                    GetLocalizedString = "粗体"
                Case efnGUIStr_FontPicker_DrawSample_Italic
                    GetLocalizedString = "斜体"
                ' FontSizeChanger
                Case efnGUIStr_FontSizeChanger_Extender_ToolTipText
                    GetLocalizedString = "单击加号或减号修改字体大小 ，当前字体大小："
                Case efnGUIStr_FontSizeChanger_mnuDefaultFontSize_Caption
                    GetLocalizedString = "设置默认值"
                Case efnGUIStr_FontSizeChanger_btnMinus_ToolTipText
                    GetLocalizedString = "减小字体大小"
                Case efnGUIStr_FontSizeChanger_btnPlus_ToolTipText
                    GetLocalizedString = "增加字体大小"
                ' FlexFn
                Case efnGUIStr_FlexFn_PrintButton_ToolTipText_Default
                    GetLocalizedString = "打印"
                Case efnGUIStr_FlexFn_PrintPreviewButton_ToolTipText_Default
                    GetLocalizedString = "打印设置和打印预览"
                Case efnGUIStr_FlexFn_FindButton_ToolTipText_Default
                    GetLocalizedString = "查找"
                Case efnGUIStr_FlexFn_CopyButton_ToolTipText_Default
                    GetLocalizedString = "复制"
                Case efnGUIStr_FlexFn_SaveButton_ToolTipText_Default
                    GetLocalizedString = "保存到文件"
                Case efnGUIStr_FlexFn_GroupDataButton_ToolTipText_Default
                    GetLocalizedString = "合并相同列"
                Case efnGUIStr_FlexFn_GroupDataButtonPressed_ToolTipText_Default
                    GetLocalizedString = "不合并相同列"
                Case efnGUIStr_FlexFn_ConfigColumnsButton_ToolTipText_Default
                    GetLocalizedString = "配置要在此报表中显示的列"
                Case efnGUIStr_FlexFn_ConfigColumnsButtonColsHidden_ToolTipText_Default
                    GetLocalizedString = "配置列（隐藏列）"
                Case efnGUIStr_FlexFn_CopyCellMenuCaption_Default
                    GetLocalizedString = "单元格"
                Case efnGUIStr_FlexFn_CopyRowMenuCaption_Default
                    GetLocalizedString = "行"
                Case efnGUIStr_FlexFn_CopyColumnMenuCaption_Default
                    GetLocalizedString = "列"
                Case efnGUIStr_FlexFn_CopyAllMenuCaption_Default
                    GetLocalizedString = "全部"
                Case efnGUIStr_FlexFn_CopySelectionMenuCaption_Default
                    GetLocalizedString = "已选择"
                Case efnGUIStr_FlexFn_mnuCopyParent_Caption
                    GetLocalizedString = "复制..."
                ' History
                Case efnGUIStr_History_mnuDelete_Caption1
                    GetLocalizedString = "删除"
                Case efnGUIStr_History_mnuDelete_Caption2
                    GetLocalizedString = "从历史纪录"
                Case efnGUIStr_History_ToolTipTextStart_Default
                    GetLocalizedString = "转到"
                Case efnGUIStr_History_ToolTipTextSelect_Default
                    GetLocalizedString = "(或者单击右键选择)"
                Case efnGUIStr_History_BackButtonToolTipText_Default
                    GetLocalizedString = "转到上一条(或者单击右键选择)"
                Case efnGUIStr_History_ForwardButtonToolTipText_Default
                    GetLocalizedString = "转到下一条(或者单击右键选择)"
                ' DateEnter
                Case efnGUIStr_DateEnter_ToolTipTextStart_Default
                    GetLocalizedString = "按格式输入日期"
                Case efnGUIStr_DateEnter_ToolTipTextEnd_Default
                    GetLocalizedString = "或单击箭头选择"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxTitle
                    GetLocalizedString = "日期输入错误"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError1
                    GetLocalizedString = "您没有在日期条目中输入天数。"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError2
                    GetLocalizedString = "天数不能为零。"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError3
                    GetLocalizedString = "不能超过31天"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError4
                    GetLocalizedString = "您没有在日期条目中输入月份。"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError5
                    GetLocalizedString = "月份的值必须在1到12之间。"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError6
                    GetLocalizedString = "您没有在日期条目中输入年份。"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError7
                    GetLocalizedString = "日期不能低于"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError8
                    GetLocalizedString = "日期不能大于"
                ' Class modules
                ' cGridHandler
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString1
                    GetLocalizedString = "按此列排序" 'Order by this column"
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString2
                    GetLocalizedString = "Order by"
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString3
                    GetLocalizedString = "正向排序" 'ascending"
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString4
                    GetLocalizedString = "反向排序" 'descending"
            
                ' FlexFnObject
                Case efnGUIStr_FlexFnObject_FindTextInGrid_InputBoxEnterTextMessageTitle
                    GetLocalizedString = "查找文字"
                Case efnGUIStr_FlexFnObject_FindTextInGrid_InputBoxEnterTextMessage
                    GetLocalizedString = "请输入要查找的文字:"
                Case efnGUIStr_FlexFnObject_FindTextInGrid_MsgboxTextNotFound
                    GetLocalizedString = "没有找到要查找的文字."
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_iDlg_Filter
                    GetLocalizedString = "文件 Excel (*.xls)|*.xls"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1a
                    GetLocalizedString = "文件"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1b
                    GetLocalizedString = "已经存在，是否覆盖？"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2a
                    GetLocalizedString = "文件不能被替换，它可以用Excel打开。要用相同的名称保存它，必须先关闭它。"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2b
                    GetLocalizedString = "是否重试？（在关闭后按“是”，“否”选择另一个名称，或“取消”取消操作"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox3
                    GetLocalizedString = "是否现在打开保存后的文件？"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox4
                    GetLocalizedString = "文件保存在"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox5
                    GetLocalizedString = "发生错误:"
                Case efnGUIStr_FlexFnObject_PageNumbersFormatStrings_Count
                    GetLocalizedString = "15"
                Case efnGUIStr_FlexFnObject_PageNumbersFormatStrings
                    Select Case nIndex
                        Case 1 ' the index starts in 1, the 0 is retrieved from efnGUIStr_General_PageNumbersFormatString_Default
                            GetLocalizedString = "页 #"
                        Case 2
                            GetLocalizedString = "页 #"
                        Case 3
                            GetLocalizedString = "页 # / N"
                        Case 4
                            GetLocalizedString = "页 # / N"
                        Case 5
                            GetLocalizedString = "页. #"
                        Case 6
                            GetLocalizedString = "# / N"
                        Case 7
                            GetLocalizedString = "页. # / N"
                        Case 8
                            GetLocalizedString = "#/N"
                        Case 9
                            GetLocalizedString = "页. #/N"
                        Case 10
                            GetLocalizedString = "页 #/N"
                        Case 11
                            GetLocalizedString = "页 #/N"
                        Case 12
                            GetLocalizedString = "# / N"
                        Case 13
                            GetLocalizedString = "页. # / N"
                        Case 14
                            GetLocalizedString = "页 # / N"
                        Case 15
                            GetLocalizedString = "页 # / N"
                    End Select
                ' cPrinterEx
                Case efnGUIStr_cPrinterEx_PrintDocumentTooLongWarning_MsgBoxWarning
                    GetLocalizedString = "文档太长，将无法完全打印。"
                ' Bas modules
                ' mGlobals
                Case efnGUIStr_mGlobals_ValidFileName_DefaultFileName
                    GetLocalizedString = "未命名"
                End Select
        Case vxLangSPANISH
            Select Case nID
                ' Forms
                ' General
                Case efnGUIStr_General_CloseButton_Caption
                    GetLocalizedString = "&Cerrar"
                Case efnGUIStr_General_OKButton_Caption
                    GetLocalizedString = "&Aceptar"
                Case efnGUIStr_General_CancelButton_Caption
                    GetLocalizedString = "&Cancelar"
                Case efnGUIStr_General_PageNumbersFormatString_Default
                    GetLocalizedString = "#"
                ' frmClipboardCopiedMessage
                Case efnGUIStr_frmClipboardCopiedMessage_lblMessage_Caption
                    GetLocalizedString = "Se copi� el texto"
                ' frmConfigHistory
                Case efnGUIStr_frmConfigHistory_Caption
                    GetLocalizedString = "Configuraci髇 de historial"
                Case efnGUIStr_frmConfigHistory_chkRememberHistory_Caption
                    GetLocalizedString = "Recordar el historial a trav閟 de sesiones"
                Case efnGUIStr_frmConfigHistory_HelpMessageTitle
                    GetLocalizedString = "Historial"
                Case efnGUIStr_frmConfigHistory_HelpMessage
                    GetLocalizedString = "Indica si el programa recordar� lo buscado o seleccionado la pr髕ima vez que lo corra, luego de cerrarlo." & vbCrLf & vbCrLf & "Puede eliminar el historial actual haciendo click en uno de los botones." & vbCrLf & "Tambi閚 se puede eliminar un solo elemento de un historial haciendo click con el bot髇 derecho del mouse sobre el mismo (en la lista, antes de entrar a esta pantalla de configuraci髇)."
                Case efnGUIStr_frmConfigHistory_cmdEraseContext_Caption
                    GetLocalizedString = "Eliminar el historial para este contexto"
                Case efnGUIStr_frmConfigHistory_cmdEraseAll_Caption
                    GetLocalizedString = "Eliminar todos los historiales"
                ' frmCopyGridTextOptions
                Case efnGUIStr_frmCopyGridTextOptions_Caption
                    GetLocalizedString = "Opciones de copia de texto"
                Case efnGUIStr_frmCopyGridTextOptions_lblColumnsSeparationMode_Caption
                    GetLocalizedString = "Separaci髇 de las columnas:"
                Case efnGUIStr_frmCopyGridTextOptions_cboMode_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "Con Tabs (mejor para copiar en programas como Excel)"
                        Case 1
                            GetLocalizedString = "Con espaciado uniforme (mejor para copiar en editores de texto con fuente de ancho fijo)"
                        Case 2
                            GetLocalizedString = "Con espaciado de acuerdo a la fuente (mejor para copiar en editores de texto con fuente de ancho variable)"
                        Case 3
                            GetLocalizedString = "Con un caracter o texto especial como separador"
                    End Select
                Case efnGUIStr_lblSelectComunsToInclude_Caption
                    GetLocalizedString = "Seleccionar qu� columnas incluir:"
                Case efnGUIStr_EnterColumnSeparatorMessage
                    GetLocalizedString = "Por favor ingrese el texto o caracter separador de columnas"
                Case efnGUIStr_EnterColumnSeparatorMessageTitle
                    GetLocalizedString = "Ingresar separador"
                Case efnGUIStr_SelectFontMessage
                    GetLocalizedString = "Necesita seleccionar qu� fuente va a usar en el programa destino donde lo va a pegar." & vbCrLf & "El alineado de las columnas ser� aproximado."
                ' frmPrintGridFormatOptions
                Case efnGUIStr_frmPrintGridFormatOptions_Caption
                    GetLocalizedString = "Formato de impresi髇"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_0
                    GetLocalizedString = "Opciones"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_1
                    GetLocalizedString = "Estilo"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_2
                    GetLocalizedString = "Otras"
                Case efnGUIStr_frmPrintGridFormatOptions_chkEnableAutoOrientation_Caption
                    GetLocalizedString = "Cambiar autom醫icamente la orientaci髇 de la p醙ina a horizontal si el listado es m醩 ancho que el papel."
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersSeparatorLine_Caption
                    GetLocalizedString = "Sep. encabezado"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintColumnsHeadersLines_Caption
                    GetLocalizedString = "Lin. Col. encabezados"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersBorder_Caption
                    GetLocalizedString = "Borde encabezados"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintFixedColsBackground_Caption
                    GetLocalizedString = "Fondo columnas fijas"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersBackground_Caption
                    GetLocalizedString = "Fondo encabezados"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintOtherBackgrounds_Caption
                    GetLocalizedString = "Otros colores de fondo"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintRowsLines_Caption
                    GetLocalizedString = "Lineas de filas"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintColumnsDataLines_Caption
                    GetLocalizedString = "Lin. columnas datos"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintOuterBorder_Caption
                    GetLocalizedString = "Borde exterior"
                Case efnGUIStr_frmPrintGridFormatOptions_lblLineWidth_Caption
                    GetLocalizedString = "Grosor de l韓eas:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblStyle_Caption
                    GetLocalizedString = "Estilo:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblOtherTextsFont_Caption
                    GetLocalizedString = "Fuente otros textos:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblSubheadingFont_Caption
                    GetLocalizedString = "Fuente del sub-encabezado:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblHeadingFont_Caption
                    GetLocalizedString = "Fuente del encabezado o t韙ulo:"
                Case efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersFont_Caption
                    GetLocalizedString = "Fuente de n鷐eros de p醙ina:"
                Case efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersFormat_Caption
                    GetLocalizedString = "Formato de n鷐eros de p醙ina:"
                Case efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersPosition_Caption
                    GetLocalizedString = "Posici髇 de n鷐eros de p醙ina:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblGridAlign_Caption
                    GetLocalizedString = "Alineaci髇 de grilla de datos:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblColor_Caption
                    GetLocalizedString = "Color:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblScalePercent_Caption
                    GetLocalizedString = "Escala:"
                Case efnGUIStr_frmPrintGridFormatOptions_cboColor_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "Predeterminado"
                        Case 1
                            GetLocalizedString = "Escala de grises"
                        Case 2
                            GetLocalizedString = "Color"
                    End Select
                Case efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_cboPageNumbersPosition_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "Abajo a la derecha"
                        Case 1
                            GetLocalizedString = "Abajo a la izquierda"
                        Case 2
                            GetLocalizedString = "Abajo centrado"
                        Case 3
                            GetLocalizedString = "Arriba a la derecha"
                        Case 4
                            GetLocalizedString = "Arriba a la izquierda"
                        Case 5
                            GetLocalizedString = "Arriba centrado"
                        Case 6
                            GetLocalizedString = "No colocar n鷐eros de p醙ina"
                    End Select
                Case efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "Centrada"
                        Case 1
                            GetLocalizedString = "Izquierda"
                        Case 2
                            GetLocalizedString = "Derecha"
                        Case 3
                            GetLocalizedString = "Estirar"
                    End Select
                Case efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_Style
                    GetLocalizedString = "Estilo"
                Case efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_CustomStyle
                    GetLocalizedString = "Estilo personal"
                Case efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_Customize
                    GetLocalizedString = "Personalizar"
                Case efnGUIStr_frmPrintGridFormatOptions_lblSample_Caption
                    GetLocalizedString = "Ejemplo:"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Column
                    GetLocalizedString = "Columna"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Element
                    GetLocalizedString = "Elemento"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Data
                    GetLocalizedString = "Dato"
                Case efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_ToolTipText
                    GetLocalizedString = "S髄o tiene efecto cuando la grilla de datos es m醩 angosta que la p醙ina"
                Case efnGUIStr_frmPrintGridFormatOptions_txtLineWidth_ToolTipText
                    GetLocalizedString = "Cambiar grosor de l韓eas (general)"
                Case efnGUIStr_frmPrintGridFormatOptions_cmdHeadersBackgroundColor_ToolTipText
                    GetLocalizedString = "Cambiar color de fondo de encabezados (y/o columnas fijas)"
                Case efnGUIStr_frmPrintGridFormatOptions_txtLineWidthHeadersSeparatorLine_ToolTipText
                    GetLocalizedString = "Grosor de l韓ea separadora de encabezados"
                Case efnGUIStr_frmPrintGridFormatOptions_VariousChangeColorCommandButtons_ToolTipText
                    GetLocalizedString = "Cambiar color"
                Case efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidth_Message
                    GetLocalizedString = "El valor del grosor de las l韓eas debe estar entre 1 y 80"
                Case efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidthHeadersSeparatorLine_Message
                    GetLocalizedString = "El valor del grosor de las l韓eas debe estar entre 1 y 80"
                ' frmPageNumbersOptions
                Case efnGUIStr_frmPageNumbersOptions_Caption
                    GetLocalizedString = "Opciones de n鷐eros de p醙ina"
                ' frmSelectColumns
                Case efnGUIStr_frmSelectColumns_Caption
                    GetLocalizedString = "Configurar columnas a ver"
                Case efnGUIStr_frmSelectColumns_lblTitle_Caption
                    GetLocalizedString = "Seleccione las columnas que desea ver:"
                Case efnGUIStr_frmSelectColumns_OneVisible_Message
                    GetLocalizedString = "Por lo menos una columna debe estar visible."
                ' frmSettingGridDataProgress
                Case efnGUIStr_frmSettingGridDataProgress_lblMessage_Caption_Start
                    GetLocalizedString = "Generando vista previa"
                Case efnGUIStr_frmSettingGridDataProgress_lblMessage_Caption_Progress
                    GetLocalizedString = "Generando vista previa, p醙ina"
                ' frmPrintPreview
                Case efnGUIStr_frmPrintPreview_Caption
                    GetLocalizedString = "Vista preliminar de impresi髇"
                Case efnGUIStr_frmPrintPreview_mnuView2p_Caption
                    GetLocalizedString = "Ver 2 p醙inas"
                Case efnGUIStr_frmPrintPreview_mnuView3p_Caption
                    GetLocalizedString = "Ver 3 p醙inas"
                Case efnGUIStr_frmPrintPreview_mnuView6p_Caption
                    GetLocalizedString = "Ver 6 p醙inas"
                Case efnGUIStr_frmPrintPreview_mnuView12p_Caption
                    GetLocalizedString = "Ver 12 p醙inas"
                Case efnGUIStr_frmPrintPreview_mnuIconsAuto_Caption
                    GetLocalizedString = "Autom醫ico"
                Case efnGUIStr_frmPrintPreview_mnuIconsSmall_Caption
                    GetLocalizedString = "蚦onos peque駉s"
                Case efnGUIStr_frmPrintPreview_mnuIconsMedium_Caption
                    GetLocalizedString = "蚦onos  medianos"
                Case efnGUIStr_frmPrintPreview_mnuIconsBig_Caption
                    GetLocalizedString = "蚦onos grandes"
                Case efnGUIStr_frmPrintPreview_mnuIconsShowBottomToolBar_Caption
                    GetLocalizedString = "Mostrar barra inferior"
                Case efnGUIStr_frmPrintPreview_CurrentlySelected
                    GetLocalizedString = "seleccionado"
                Case efnGUIStr_frmPrintPreview_CurrentlyShown
                    GetLocalizedString = "actualmente visible"
                Case efnGUIStr_frmPrintPreview_lblPageOrientation_Caption
                    GetLocalizedString = "Orientaci髇 de p醙ina:"
                Case efnGUIStr_frmPrintPreview_lblView_Caption
                    GetLocalizedString = "Ver:"
                Case efnGUIStr_frmPrintPreview_lblScalePercent_Caption
                    GetLocalizedString = "Escala:"
                Case efnGUIStr_frmPrintPreview_lblPage_Singular_Caption
                    GetLocalizedString = "P醙ina:"
                Case efnGUIStr_frmPrintPreview_lblPage_Plural_Caption
                    GetLocalizedString = "P醙inas:"
                Case efnGUIStr_frmPrintPreview_lblPageCount_Caption
                    GetLocalizedString = "de"
                Case efnGUIStr_frmPrintPreview_PreparingDoc_Caption
                    GetLocalizedString = "Generando vista preliminar..."
                Case efnGUIStr_frmPrintPreview_cmdClose_Caption
                    GetLocalizedString = "Cerrar vista preliminar"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Print
                    GetLocalizedString = "Imprimir"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageSetup
                    GetLocalizedString = "Configurar p醙ina"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageNumbersOptions
                    GetLocalizedString = "Opciones de n鷐eros de p醙ina"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Format
                    GetLocalizedString = "Formato"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationPortrait
                    GetLocalizedString = "Orientaci髇 vertical"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationLandscape
                    GetLocalizedString = "Orientaci髇 horizontal"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewNormalSize
                    GetLocalizedString = "Ver tama駉 de p醙ina normal"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenWidth
                    GetLocalizedString = "Ver p醙ina ajustada al ancho de la pantalla"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenHeight
                    GetLocalizedString = "Ver p醙ina ajustada al alto de la pantalla"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewSeveralPages
                    GetLocalizedString = "Ver varias p醙inas a la vez"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_DecreaseScale
                    GetLocalizedString = "Disminuir tama駉 de fuentes y elementos"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_IncreaseScale
                    GetLocalizedString = "Aumentar tama駉 de fuentes y elementos"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_FirstPage
                    GetLocalizedString = "Primera p醙ina"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Singular
                    GetLocalizedString = "P醙ina anterior"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Plural
                    GetLocalizedString = "P醙inas anteriores"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Singular
                    GetLocalizedString = "P醙ina siguiente"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Plural
                    GetLocalizedString = "Pr髕imas p醙inas"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_LastPage
                    GetLocalizedString = "趌tima p醙ina"
                
                ' UserControls
                ' FontPicker
                Case efnGUIStr_FontPicker_ButtonToolTipTextDefault
                    GetLocalizedString = "Seleccionar fuente"
                Case efnGUIStr_FontPicker_DrawSample_Bold
                    GetLocalizedString = "negrita"
                Case efnGUIStr_FontPicker_DrawSample_Italic
                    GetLocalizedString = "cursiva"
                ' FontSizeChanger
                Case efnGUIStr_FontSizeChanger_Extender_ToolTipText
                    GetLocalizedString = "Haga clic en los signos + y - si desea cambiar el tama駉 de la letra (el tama駉 actual es "
                Case efnGUIStr_FontSizeChanger_mnuDefaultFontSize_Caption
                    GetLocalizedString = "Poner valor por defecto"
                Case efnGUIStr_FontSizeChanger_btnMinus_ToolTipText
                    GetLocalizedString = "Disminuir el tama駉 de la letra"
                Case efnGUIStr_FontSizeChanger_btnPlus_ToolTipText
                    GetLocalizedString = "Aumentar el tama駉 de la letra"
                ' FlexFn
                Case efnGUIStr_FlexFn_PrintButton_ToolTipText_Default
                    GetLocalizedString = "Imprimir"
                Case efnGUIStr_FlexFn_PrintPreviewButton_ToolTipText_Default
                    GetLocalizedString = "Configuraci髇 de impresi髇 y vista preliminar"
                Case efnGUIStr_FlexFn_FindButton_ToolTipText_Default
                    GetLocalizedString = "Buscar"
                Case efnGUIStr_FlexFn_CopyButton_ToolTipText_Default
                    GetLocalizedString = "Copiar"
                Case efnGUIStr_FlexFn_SaveButton_ToolTipText_Default
                    GetLocalizedString = "Guardar en un archivo"
                Case efnGUIStr_FlexFn_GroupDataButton_ToolTipText_Default
                    GetLocalizedString = "Agrupar textos que son iguales en las columnas"
                Case efnGUIStr_FlexFn_GroupDataButtonPressed_ToolTipText_Default
                    GetLocalizedString = "No agrupar textos en las columnas"
                Case efnGUIStr_FlexFn_ConfigColumnsButton_ToolTipText_Default
                    GetLocalizedString = "Configurar qu� columnas mostrar en este listado"
                Case efnGUIStr_FlexFn_ConfigColumnsButtonColsHidden_ToolTipText_Default
                    GetLocalizedString = "Configurar columnas (hay columnas ocultas)"
                Case efnGUIStr_FlexFn_CopyCellMenuCaption_Default
                    GetLocalizedString = "Celda"
                Case efnGUIStr_FlexFn_CopyRowMenuCaption_Default
                    GetLocalizedString = "L韓ea"
                Case efnGUIStr_FlexFn_CopyColumnMenuCaption_Default
                    GetLocalizedString = "Columna"
                Case efnGUIStr_FlexFn_CopyAllMenuCaption_Default
                    GetLocalizedString = "Todo"
                Case efnGUIStr_FlexFn_CopySelectionMenuCaption_Default
                    GetLocalizedString = "Selecci髇"
                Case efnGUIStr_FlexFn_mnuCopyParent_Caption
                    GetLocalizedString = "Copiar..."
                ' History
                Case efnGUIStr_History_mnuDelete_Caption1
                    GetLocalizedString = "Eliminar"
                Case efnGUIStr_History_mnuDelete_Caption2
                    GetLocalizedString = "del historial"
                Case efnGUIStr_History_ToolTipTextStart_Default
                    GetLocalizedString = "Ir a"
                Case efnGUIStr_History_ToolTipTextSelect_Default
                    GetLocalizedString = "(o clic con el bot髇 derecho para seleccionar)"
                Case efnGUIStr_History_BackButtonToolTipText_Default
                    GetLocalizedString = "Ir a 韙em anterior (o click con el bot髇 derecho para seleccionar)"
                Case efnGUIStr_History_ForwardButtonToolTipText_Default
                    GetLocalizedString = "Ir a 韙em siguiente (o click con el bot髇 derecho para seleccionar)"
                ' DateEnter
                Case efnGUIStr_DateEnter_ToolTipTextStart_Default
                    GetLocalizedString = "Ingrese la fecha en formato"
                Case efnGUIStr_DateEnter_ToolTipTextEnd_Default
                    GetLocalizedString = "o haga clic en la flecha para seleccionar"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxTitle
                    GetLocalizedString = "Error en ingreso de fecha"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError1
                    GetLocalizedString = "No ingres� el d韆 en el ingreso de fecha."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError2
                    GetLocalizedString = "El d韆 no puede ser cero."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError3
                    GetLocalizedString = "El d韆 no puede ser mayor a 31."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError4
                    GetLocalizedString = "No ingres� el mes en el ingreso de fecha."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError5
                    GetLocalizedString = "El valor del mes debe estar entre 1 y 12."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError6
                    GetLocalizedString = "No ingres� el a駉 en el ingreso de fecha."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError7
                    GetLocalizedString = "La fecha no puede ser menor que"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError8
                    GetLocalizedString = "La fecha no puede ser mayor que"
                    
                ' Class modules
                ' cGridHandler
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString1
                    GetLocalizedString = "Ordenar por esta columna"
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString2
                    GetLocalizedString = "Ordenar por"
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString3
                    GetLocalizedString = "ascendente"
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString4
                    GetLocalizedString = "descendente"
            
                ' FlexFnObject
                Case efnGUIStr_FlexFnObject_FindTextInGrid_InputBoxEnterTextMessageTitle
                    GetLocalizedString = "Buscar texto"
                Case efnGUIStr_FlexFnObject_FindTextInGrid_InputBoxEnterTextMessage
                    GetLocalizedString = "Por favor ingrese el texto a buscar:"
                Case efnGUIStr_FlexFnObject_FindTextInGrid_MsgboxTextNotFound
                    GetLocalizedString = "Texto no encontrado."
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_iDlg_Filter
                    GetLocalizedString = "Archivos de Excel (*.xls)|*.xls"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1a
                    GetLocalizedString = "El archivo"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1b
                    GetLocalizedString = "ya existe, 縟esea sobreescribirlo?"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2a
                    GetLocalizedString = "El archivo no se puede reemplazar, es posible que est� abierto con Excel, para guardarlo con el mismo nombre debe cerrarlo antes."
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2b
                    GetLocalizedString = "縍eintentar? (Presione 'S�' luego de cerrarlo, 'No' para elegir otro nombre, o 'Cancelar' para cancelar la operaci髇."
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox3
                    GetLocalizedString = "緿esea abrir ahora con Excel el archivo guardado?"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox4
                    GetLocalizedString = "El archivo se guard� en"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox5
                    GetLocalizedString = "Se produjo un error:"
                Case efnGUIStr_FlexFnObject_PageNumbersFormatStrings_Count
                    GetLocalizedString = "15"
                Case efnGUIStr_FlexFnObject_PageNumbersFormatStrings
                    Select Case nIndex
                        Case 1 ' the index starts in 1, the 0 is retrieved from efnGUIStr_General_PageNumbersFormatString_Default
                            GetLocalizedString = "P醙ina #"
                        Case 2
                            GetLocalizedString = "p醙ina #"
                        Case 3
                            GetLocalizedString = "P醙ina # de N"
                        Case 4
                            GetLocalizedString = "p醙ina # de N"
                        Case 5
                            GetLocalizedString = "P醙. #"
                        Case 6
                            GetLocalizedString = "# de N"
                        Case 7
                            GetLocalizedString = "P醙. # de N"
                        Case 8
                            GetLocalizedString = "#/N"
                        Case 9
                            GetLocalizedString = "P醙. #/N"
                        Case 10
                            GetLocalizedString = "P醙ina #/N"
                        Case 11
                            GetLocalizedString = "p醙ina #/N"
                        Case 12
                            GetLocalizedString = "# / N"
                        Case 13
                            GetLocalizedString = "P醙. # / N"
                        Case 14
                            GetLocalizedString = "P醙ina # / N"
                        Case 15
                            GetLocalizedString = "p醙ina # / N"
                    End Select
                ' cPrinterEx
                Case efnGUIStr_cPrinterEx_PrintDocumentTooLongWarning_MsgBoxWarning
                    GetLocalizedString = "Documento demasiado extenso, no se imprimir� completo."
                ' Bas modules
                ' mGlobals
                Case efnGUIStr_mGlobals_ValidFileName_DefaultFileName
                    GetLocalizedString = "Sin t韙ulo"
            End Select
        
        Case Else ' ENGLISH
            Select Case nID
                ' General
                Case efnGUIStr_General_CloseButton_Caption
                    GetLocalizedString = "&Close"
                Case efnGUIStr_General_OKButton_Caption
                    GetLocalizedString = "&OK"
                Case efnGUIStr_General_CancelButton_Caption
                    GetLocalizedString = "&Cancel"
                Case efnGUIStr_General_PageNumbersFormatString_Default
                    GetLocalizedString = "#"
                ' frmClipboardCopiedMessage
                Case efnGUIStr_frmClipboardCopiedMessage_lblMessage_Caption
                    GetLocalizedString = "Text copied"
                ' frmConfigHistory
                Case efnGUIStr_frmConfigHistory_Caption
                    GetLocalizedString = "History configuration"
                Case efnGUIStr_frmConfigHistory_chkRememberHistory_Caption
                    GetLocalizedString = "Remember the history across sessions"
                Case efnGUIStr_frmConfigHistory_HelpMessageTitle
                    GetLocalizedString = "History"
                Case efnGUIStr_frmConfigHistory_HelpMessage
                    GetLocalizedString = "Determines if the program will remember what was searched or selected the next times that it runs, after closing it." & vbCrLf & vbCrLf & "You can erase the current history by clicking one of the buttons." & vbCrLf & "You can also erase only one item of the history by clicking with the right mouse's button on it (in the list, before entering this configuration screen)."
                Case efnGUIStr_frmConfigHistory_cmdEraseContext_Caption
                    GetLocalizedString = "Erase history for this context"
                Case efnGUIStr_frmConfigHistory_cmdEraseAll_Caption
                    GetLocalizedString = "Erase all"
                ' frmCopyGridTextOptions
                Case efnGUIStr_frmCopyGridTextOptions_Caption
                    GetLocalizedString = "Copy text options"
                Case efnGUIStr_frmCopyGridTextOptions_lblColumnsSeparationMode_Caption
                    GetLocalizedString = "Separation of the columns:"
                Case efnGUIStr_frmCopyGridTextOptions_cboMode_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "With Tabs (best to copy to programs like Excel)"
                        Case 1
                            GetLocalizedString = "With uniform spacing (best to copy to text editors with fixed width fonts)"
                        Case 2
                            GetLocalizedString = "With spacing according to font (best to copy to text editors with variable width fonts)"
                        Case 3
                            GetLocalizedString = "With a custom character or text as the separator"
                    End Select
                Case efnGUIStr_lblSelectComunsToInclude_Caption
                    GetLocalizedString = "Select columns to include:"
                Case efnGUIStr_EnterColumnSeparatorMessage
                    GetLocalizedString = "Please enter the text or character as the column separator"
                Case efnGUIStr_EnterColumnSeparatorMessageTitle
                    GetLocalizedString = "Enter separator"
                Case efnGUIStr_SelectFontMessage
                    GetLocalizedString = "Please select what font is going to use the destination program where you are going to paste it." & vbCrLf & "The alignment of the columns will be approximate."
                ' frmPrintGridFormatOptions
                Case efnGUIStr_frmPrintGridFormatOptions_Caption
                    GetLocalizedString = "Printing format"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_0
                    GetLocalizedString = "Options"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_1
                    GetLocalizedString = "Style"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_2
                    GetLocalizedString = "More"
                Case efnGUIStr_frmPrintGridFormatOptions_chkEnableAutoOrientation_Caption
                    GetLocalizedString = "Automatically change the page orientation to horizontal if the report is wider than the paper."
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersSeparatorLine_Caption
                    GetLocalizedString = "Headers Sep."
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintColumnsHeadersLines_Caption
                    GetLocalizedString = "Column headers lines"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersBorder_Caption
                    GetLocalizedString = "Headers borders"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintFixedColsBackground_Caption
                    GetLocalizedString = "Fixed columns Bckgr."
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersBackground_Caption
                    GetLocalizedString = "Headers background"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintOtherBackgrounds_Caption
                    GetLocalizedString = "Other background colors"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintRowsLines_Caption
                    GetLocalizedString = "Row lines"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintColumnsDataLines_Caption
                    GetLocalizedString = "Columns data lines"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintOuterBorder_Caption
                    GetLocalizedString = "Outer edge"
                Case efnGUIStr_frmPrintGridFormatOptions_lblLineWidth_Caption
                    GetLocalizedString = "Lines width:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblStyle_Caption
                    GetLocalizedString = "Style:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblOtherTextsFont_Caption
                    GetLocalizedString = "Other texts font:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblSubheadingFont_Caption
                    GetLocalizedString = "Sub-heading font:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblHeadingFont_Caption
                    GetLocalizedString = "Heading or title font:"
                Case efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersFont_Caption
                    GetLocalizedString = "Page numbers font:"
                Case efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersFormat_Caption
                    GetLocalizedString = "Page numbers format:"
                Case efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersPosition_Caption
                    GetLocalizedString = "Page numbers position:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblGridAlign_Caption
                    GetLocalizedString = "Grid alignment:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblColor_Caption
                    GetLocalizedString = "Color:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblScalePercent_Caption
                    GetLocalizedString = "Scale:"
                Case efnGUIStr_frmPrintGridFormatOptions_cboColor_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "Printer default"
                        Case 1
                            GetLocalizedString = "Grey scale"
                        Case 2
                            GetLocalizedString = "Color"
                    End Select
                Case efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_cboPageNumbersPosition_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "Bottom right"
                        Case 1
                            GetLocalizedString = "Bottom left"
                        Case 2
                            GetLocalizedString = "Bottom centered"
                        Case 3
                            GetLocalizedString = "Top right"
                        Case 4
                            GetLocalizedString = "Top left"
                        Case 5
                            GetLocalizedString = "Top centered"
                        Case 6
                            GetLocalizedString = "Don't add page numbers"
                    End Select
                Case efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "Centered"
                        Case 1
                            GetLocalizedString = "Left"
                        Case 2
                            GetLocalizedString = "Right"
                        Case 3
                            GetLocalizedString = "Stretch"
                    End Select
                Case efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_Style
                    GetLocalizedString = "Style"
                Case efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_CustomStyle
                    GetLocalizedString = "Custom style"
                Case efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_Customize
                    GetLocalizedString = "Customize"
                Case efnGUIStr_frmPrintGridFormatOptions_lblSample_Caption
                    GetLocalizedString = "Sample:"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Column
                    GetLocalizedString = "Column"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Element
                    GetLocalizedString = "Element"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Data
                    GetLocalizedString = "Data"
                Case efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_ToolTipText
                    GetLocalizedString = "It only has effect when the data grid is narrower than the page"
                Case efnGUIStr_frmPrintGridFormatOptions_txtLineWidth_ToolTipText
                    GetLocalizedString = "Change line thickness (general)"
                Case efnGUIStr_frmPrintGridFormatOptions_cmdHeadersBackgroundColor_ToolTipText
                    GetLocalizedString = "Change background color of headers (and / or fixed columns)"
                Case efnGUIStr_frmPrintGridFormatOptions_txtLineWidthHeadersSeparatorLine_ToolTipText
                    GetLocalizedString = "Headers separator line thickness"
                Case efnGUIStr_frmPrintGridFormatOptions_VariousChangeColorCommandButtons_ToolTipText
                    GetLocalizedString = "Change color"
                Case efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidth_Message
                    GetLocalizedString = "The thickness value of the lines must be between 1 and 80"
                Case efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidthHeadersSeparatorLine_Message
                    GetLocalizedString = "The thickness value of the lines must be between 1 and 80"
                ' frmPageNumbersOptions
                Case efnGUIStr_frmPageNumbersOptions_Caption
                    GetLocalizedString = "Page numbers options"
                ' frmSelectColumns
                Case efnGUIStr_frmSelectColumns_Caption
                    GetLocalizedString = "Configure visible columns"
                Case efnGUIStr_frmSelectColumns_lblTitle_Caption
                    GetLocalizedString = "Select the columns to display:"
                Case efnGUIStr_frmSelectColumns_OneVisible_Message
                    GetLocalizedString = "At least one column must be selected."
                ' frmSettingGridDataProgress
                Case efnGUIStr_frmSettingGridDataProgress_lblMessage_Caption_Start
                    GetLocalizedString = "Generating preview"
                Case efnGUIStr_frmSettingGridDataProgress_lblMessage_Caption_Progress
                    GetLocalizedString = "Generating preview, page"
                ' frmPrintPreview
                Case efnGUIStr_frmPrintPreview_Caption
                    GetLocalizedString = "Print preview"
                Case efnGUIStr_frmPrintPreview_mnuView2p_Caption
                    GetLocalizedString = "View 2 pages"
                Case efnGUIStr_frmPrintPreview_mnuView3p_Caption
                    GetLocalizedString = "View 3 pages"
                Case efnGUIStr_frmPrintPreview_mnuView6p_Caption
                    GetLocalizedString = "View 6 pages"
                Case efnGUIStr_frmPrintPreview_mnuView12p_Caption
                    GetLocalizedString = "Ver 12 p醙inas"
                Case efnGUIStr_frmPrintPreview_mnuIconsAuto_Caption
                    GetLocalizedString = "Automatic"
                Case efnGUIStr_frmPrintPreview_mnuIconsSmall_Caption
                    GetLocalizedString = "Small icons"
                Case efnGUIStr_frmPrintPreview_mnuIconsMedium_Caption
                    GetLocalizedString = "Medium icons"
                Case efnGUIStr_frmPrintPreview_mnuIconsBig_Caption
                    GetLocalizedString = "Large icons"
                Case efnGUIStr_frmPrintPreview_mnuIconsShowBottomToolBar_Caption
                    GetLocalizedString = "Show bottom bar"
                Case efnGUIStr_frmPrintPreview_CurrentlySelected
                    GetLocalizedString = "selected"
                Case efnGUIStr_frmPrintPreview_CurrentlyShown
                    GetLocalizedString = "now shown"
                Case efnGUIStr_frmPrintPreview_lblPageOrientation_Caption
                    GetLocalizedString = "Page orientation:"
                Case efnGUIStr_frmPrintPreview_lblView_Caption
                    GetLocalizedString = "View:"
                Case efnGUIStr_frmPrintPreview_lblScalePercent_Caption
                    GetLocalizedString = "Scale:"
                Case efnGUIStr_frmPrintPreview_lblPage_Singular_Caption
                    GetLocalizedString = "Page:"
                Case efnGUIStr_frmPrintPreview_lblPage_Plural_Caption
                    GetLocalizedString = "Pages:"
                Case efnGUIStr_frmPrintPreview_lblPageCount_Caption
                    GetLocalizedString = "of"
                Case efnGUIStr_frmPrintPreview_PreparingDoc_Caption
                    GetLocalizedString = "Generating print preview..."
                Case efnGUIStr_frmPrintPreview_cmdClose_Caption
                    GetLocalizedString = "Close print preview"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Print
                    GetLocalizedString = "Print"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageSetup
                    GetLocalizedString = "Page setup"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageNumbersOptions
                    GetLocalizedString = "Page numbers options"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Format
                    GetLocalizedString = "Format"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationPortrait
                    GetLocalizedString = "Orientation portrait"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationLandscape
                    GetLocalizedString = "Orientation landscape"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewNormalSize
                    GetLocalizedString = "View normal page size"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenWidth
                    GetLocalizedString = "View page adjusted to the screen width"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenHeight
                    GetLocalizedString = "View page adjusted to the screen height"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewSeveralPages
                    GetLocalizedString = "View several pages"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_DecreaseScale
                    GetLocalizedString = "Decrease fonts and elements size"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_IncreaseScale
                    GetLocalizedString = "Increase fonts and elements size"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_FirstPage
                    GetLocalizedString = "First page"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Singular
                    GetLocalizedString = "Previous page"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Plural
                    GetLocalizedString = "Previous pages"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Singular
                    GetLocalizedString = "Next page"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Plural
                    GetLocalizedString = "Next pages"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_LastPage
                    GetLocalizedString = "Last page"
                
                ' UserControls
                ' FontPicker
                Case efnGUIStr_FontPicker_ButtonToolTipTextDefault
                    GetLocalizedString = "Select font"
                Case efnGUIStr_FontPicker_DrawSample_Bold
                    GetLocalizedString = "bold"
                Case efnGUIStr_FontPicker_DrawSample_Italic
                    GetLocalizedString = "italic"
                ' FontSizeChanger
                Case efnGUIStr_FontSizeChanger_Extender_ToolTipText
                    GetLocalizedString = "Click on the + and - signs if you want to change the font size (the current size is "
                Case efnGUIStr_FontSizeChanger_mnuDefaultFontSize_Caption
                    GetLocalizedString = "Set default value"
                Case efnGUIStr_FontSizeChanger_btnMinus_ToolTipText
                    GetLocalizedString = "Decrease font size"
                Case efnGUIStr_FontSizeChanger_btnPlus_ToolTipText
                    GetLocalizedString = "Increase font size"
                ' FlexFn
                Case efnGUIStr_FlexFn_PrintButton_ToolTipText_Default
                    GetLocalizedString = "Print"
                Case efnGUIStr_FlexFn_PrintPreviewButton_ToolTipText_Default
                    GetLocalizedString = "Print settings and print preview"
                Case efnGUIStr_FlexFn_FindButton_ToolTipText_Default
                    GetLocalizedString = "Find"
                Case efnGUIStr_FlexFn_CopyButton_ToolTipText_Default
                    GetLocalizedString = "Copy"
                Case efnGUIStr_FlexFn_SaveButton_ToolTipText_Default
                    GetLocalizedString = "Save to a file"
                Case efnGUIStr_FlexFn_GroupDataButton_ToolTipText_Default
                    GetLocalizedString = "Group texts that are the same in columns"
                Case efnGUIStr_FlexFn_GroupDataButtonPressed_ToolTipText_Default
                    GetLocalizedString = "Do not group texts in columns"
                Case efnGUIStr_FlexFn_ConfigColumnsButton_ToolTipText_Default
                    GetLocalizedString = "Configure what columns to show in this report"
                Case efnGUIStr_FlexFn_ConfigColumnsButtonColsHidden_ToolTipText_Default
                    GetLocalizedString = "Configure columns (there are hidden columns)"
                Case efnGUIStr_FlexFn_CopyCellMenuCaption_Default
                    GetLocalizedString = "Cell"
                Case efnGUIStr_FlexFn_CopyRowMenuCaption_Default
                    GetLocalizedString = "Row"
                Case efnGUIStr_FlexFn_CopyColumnMenuCaption_Default
                    GetLocalizedString = "Column"
                Case efnGUIStr_FlexFn_CopyAllMenuCaption_Default
                    GetLocalizedString = "All"
                Case efnGUIStr_FlexFn_CopySelectionMenuCaption_Default
                    GetLocalizedString = "Selection"
                Case efnGUIStr_FlexFn_mnuCopyParent_Caption
                    GetLocalizedString = "Copy..."
                ' History
                Case efnGUIStr_History_mnuDelete_Caption1
                    GetLocalizedString = "Delete"
                Case efnGUIStr_History_mnuDelete_Caption2
                    GetLocalizedString = "from history"
                Case efnGUIStr_History_ToolTipTextStart_Default
                    GetLocalizedString = "Go to"
                Case efnGUIStr_History_ToolTipTextSelect_Default
                    GetLocalizedString = "(or click with the right button to select)"
                Case efnGUIStr_History_BackButtonToolTipText_Default
                    GetLocalizedString = "Go to previous item (or click with the right button to select)"
                Case efnGUIStr_History_ForwardButtonToolTipText_Default
                    GetLocalizedString = "Go to next item (or click with the right button to select)"
                ' DateEnter
                Case efnGUIStr_DateEnter_ToolTipTextStart_Default
                    GetLocalizedString = "Enter the date in the format"
                Case efnGUIStr_DateEnter_ToolTipTextEnd_Default
                    GetLocalizedString = "or click in the arrow to select"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxTitle
                    GetLocalizedString = "Date enter error"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError1
                    GetLocalizedString = "You did not enter the day in the date entry."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError2
                    GetLocalizedString = "The day can't be zero."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError3
                    GetLocalizedString = "The day can't be greater than 31."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError4
                    GetLocalizedString = "You did not enter the month in the date entry."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError5
                    GetLocalizedString = "The value of the month must be between 1 y 12."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError6
                    GetLocalizedString = "You did not enter the year in the date entry."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError7
                    GetLocalizedString = "The date can not be less than"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError8
                    GetLocalizedString = "The date can not be greater than"
                
                ' Class modules
                ' cGridHandler
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString1
                    GetLocalizedString = "Order by this column"
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString2
                    GetLocalizedString = "Order by"
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString3
                    GetLocalizedString = "ascending"
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString4
                    GetLocalizedString = "descending"
            
                ' FlexFnObject
                Case efnGUIStr_FlexFnObject_FindTextInGrid_InputBoxEnterTextMessageTitle
                    GetLocalizedString = "Find text"
                Case efnGUIStr_FlexFnObject_FindTextInGrid_InputBoxEnterTextMessage
                    GetLocalizedString = "Please enter the text to search for:"
                Case efnGUIStr_FlexFnObject_FindTextInGrid_MsgboxTextNotFound
                    GetLocalizedString = "Text not found."
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_iDlg_Filter
                    GetLocalizedString = "Excel files (*.xls)|*.xls"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1a
                    GetLocalizedString = "The file"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1b
                    GetLocalizedString = "already exists, do you want to overwrite it?"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2a
                    GetLocalizedString = "The file can not be replaced, it may be open with Excel. To save it with the same name, you must close it first."
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2b
                    GetLocalizedString = "Retry? (Press 'Yes' after closing it, 'No' to choose another name, or 'Cancel' to cancel the operation."
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox3
                    GetLocalizedString = "Do you want to open the saved file now with Excel?"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox4
                    GetLocalizedString = "The file was saved in"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox5
                    GetLocalizedString = "There was an error:"
                Case efnGUIStr_FlexFnObject_PageNumbersFormatStrings_Count
                    GetLocalizedString = "15"
                Case efnGUIStr_FlexFnObject_PageNumbersFormatStrings
                    Select Case nIndex
                        Case 1 ' the index starts in 1, the 0 is retrieved from efnGUIStr_General_PageNumbersFormatString_Default
                            GetLocalizedString = "Page #"
                        Case 2
                            GetLocalizedString = "page #"
                        Case 3
                            GetLocalizedString = "Page # of N"
                        Case 4
                            GetLocalizedString = "page # of N"
                        Case 5
                            GetLocalizedString = "Pg. #"
                        Case 6
                            GetLocalizedString = "# of N"
                        Case 7
                            GetLocalizedString = "Pg. # of N"
                        Case 8
                            GetLocalizedString = "#/N"
                        Case 9
                            GetLocalizedString = "Pg. #/N"
                        Case 10
                            GetLocalizedString = "Page #/N"
                        Case 11
                            GetLocalizedString = "page #/N"
                        Case 12
                            GetLocalizedString = "# / N"
                        Case 13
                            GetLocalizedString = "Pg. # / N"
                        Case 14
                            GetLocalizedString = "Page # / N"
                        Case 15
                            GetLocalizedString = "page # / N"
                    End Select
                ' cPrinterEx
                Case efnGUIStr_cPrinterEx_PrintDocumentTooLongWarning_MsgBoxWarning
                    GetLocalizedString = "Document too long, will not be printed completely."
                ' Bas modules
                ' mGlobals
                Case efnGUIStr_mGlobals_ValidFileName_DefaultFileName
                    GetLocalizedString = "Untitled"
            End Select
    End Select
        
End Function

Public Property Get GUILanguage() As vbExGUILanguageConstants
    If mGUILanguage = vxLangAUTO Then SetGUILanguage
    GUILanguage = mGUILanguage
End Property

Public Property Let GUILanguage(nLang As vbExGUILanguageConstants)
    Dim iPrev As Long
    
    If nLang <> mGUILanguage Then
        iPrev = mGUILanguage
        mGUILanguage = nLang
        ResetCommonButtonsAccelerators
        BroadcastUILanguageChange iPrev
    End If
End Property

Private Sub SetGUILanguage()
    mGUILanguage = CLng(GetUserDefaultUILanguage And &HFF)
    If Not GUILaguageIsSupported(mGUILanguage) Then
        mGUILanguage = vxLangENGLISH
    End If
End Sub

Private Function GUILaguageIsSupported(nLang As Long) As Boolean
        Select Case nLang
            Case vxLangENGLISH, vxLangSPANISH, vxLangCHINESE_SIMPLIFIED
                GUILaguageIsSupported = True
        End Select
End Function





