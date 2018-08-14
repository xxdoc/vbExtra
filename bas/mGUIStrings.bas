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
    efnGUIStr_frmPrintGridFormatOptions_lblPageNumbersFont_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblPageNumbersFormat_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblPageNumbersPosition_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblGridAlign_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblColor_Caption
    efnGUIStr_frmPrintGridFormatOptions_lblScalePercent_Caption
    efnGUIStr_frmPrintGridFormatOptions_cboColor_List
    efnGUIStr_frmPrintGridFormatOptions_cboPageNumbersPosition_List
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
    efnGUIStr_frmPrintPreview_lblPageOrientation_Caption
    efnGUIStr_frmPrintPreview_lblView_Caption
    efnGUIStr_frmPrintPreview_lblScalePercent_Caption
    efnGUIStr_frmPrintPreview_lblPage_Singular_Caption
    efnGUIStr_frmPrintPreview_lblPage_Plural_Caption
    efnGUIStr_frmPrintPreview_lblPageCount_Caption
    efnGUIStr_frmPrintPreview_lblPreparingDoc_Caption
    efnGUIStr_frmPrintPreview_cmdClose_Caption
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Print
    efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageSetup
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
                    GetLocalizedString = "&C ¹Ø±Õ"
                Case efnGUIStr_General_OKButton_Caption
                    GetLocalizedString = "&O È·¶¨"
                Case efnGUIStr_General_CancelButton_Caption
                    GetLocalizedString = "&C È¡Ïû"
                Case efnGUIStr_General_PageNumbersFormatString_Default
                    GetLocalizedString = "#"
                ' frmClipboardCopiedMessage
                Case efnGUIStr_frmClipboardCopiedMessage_lblMessage_Caption
                    GetLocalizedString = "¸´ÖÆÎÄ×Ö"
                ' frmConfigHistory
                Case efnGUIStr_frmConfigHistory_Caption
                    GetLocalizedString = "ÀúÊ·ÅäÖÃ"
                Case efnGUIStr_frmConfigHistory_chkRememberHistory_Caption
                    GetLocalizedString = "¼Ç×¡¿ç»á»°µÄÀúÊ·"
                Case efnGUIStr_frmConfigHistory_HelpMessageTitle
                    GetLocalizedString = "ÀúÊ·"
                Case efnGUIStr_frmConfigHistory_HelpMessage
                    GetLocalizedString = "È·¶¨³ÌÐòÊÇ·ñ»á¼Ç×¡ËüÏÂÒ»´Î±»ËÑË÷»òÑ¡ÔñµÄÄÚÈÝ£¬ÔÚ¹Ø±ÕËüÖ®ºó£¬Äú¿ÉÒÔÍ¨¹ýµ¥»÷ÆäÖÐÒ»¸ö°´Å¥À´É¾³ýµ±Ç°ÀúÊ·¡£" & vbCrLf & vbCrLf & "ÄúÒ²¿ÉÒÔÍ¨¹ýµ¥»÷ÓÒ¼üÀ´½öÉ¾³ýÀúÊ·¼ÇÂ¼ÖÐµÄÒ»¸öÏîÄ¿¡£ Êó±ê°´Å¥£¨ÔÚÁÐ±íÖÐ£¬ÔÚ½øÈë´ËÅäÖÃÆÁÄ»Ö®Ç°£©"
                Case efnGUIStr_frmConfigHistory_cmdEraseContext_Caption
                    GetLocalizedString = "É¾³ý´ËÉÏÏÂÎÄµÄÀúÊ·¼ÇÂ¼"
                Case efnGUIStr_frmConfigHistory_cmdEraseAll_Caption
                    GetLocalizedString = "É¾³ýÈ«²¿"
                ' frmCopyGridTextOptions
                Case efnGUIStr_frmCopyGridTextOptions_Caption
                    GetLocalizedString = "¸´ÖÆÎÄ×ÖÑ¡Ïî"
                Case efnGUIStr_frmCopyGridTextOptions_lblColumnsSeparationMode_Caption
                    GetLocalizedString = "ÁÐ·Ö¸ô·û:"
                Case efnGUIStr_frmCopyGridTextOptions_cboMode_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "Tab ¼ü£¨½¨ÒéÓÃÓÚ¸´ÖÆµ½ Excel ³ÌÐòÊ¹ÓÃ£©"
                        Case 1
                            GetLocalizedString = "ÏàÍ¬¿í¶È¿Õ¸ñ£¨×îºÃ¸´ÖÆ¾ßÓÐ¹Ì¶¨¿í¶È×ÖÌåµÄÎÄ±¾±à¼­Æ÷£©"
                        Case 2
                            GetLocalizedString = "¸ù¾Ý×ÖÌåµÄ¼ä¾à£¨×îºÃ¸´ÖÆµ½´øÓÐ¿É±ä¿í¶È×ÖÌåµÄÎÄ±¾±à¼­Æ÷£©"
                        Case 3
                            GetLocalizedString = "Ê¹ÓÃ×Ô¶¨Òå×Ö·û»òÎÄ±¾×÷Îª·Ö¸ô·û"
                    End Select
                Case efnGUIStr_lblSelectComunsToInclude_Caption
                    GetLocalizedString = "Ñ¡ÔñÒª°üº¬µÄÁÐ£º"
                Case efnGUIStr_EnterColumnSeparatorMessage
                    GetLocalizedString = "ÇëÊäÈëÎÄ±¾»ò×Ö·û×÷ÎªÁÐ·Ö¸ô·û"
                Case efnGUIStr_EnterColumnSeparatorMessageTitle
                    GetLocalizedString = "ÊäÈë·Ö¸ô·û"
                Case efnGUIStr_SelectFontMessage
                    GetLocalizedString = "ÇëÑ¡Ôñ½«ÒªÊ¹ÓÃµÄÄ¿±ê³ÌÐòÊ¹ÓÃÄÄÖÖ×ÖÌå½øÐÐÕ³Ìù,ÁÐµÄ¶ÔÆë·½Ê½Îª½üËÆÖµ."
                ' frmPrintGridFormatOptions
                Case efnGUIStr_frmPrintGridFormatOptions_Caption
                    GetLocalizedString = "´òÓ¡¸ñÊ½"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_0
                    GetLocalizedString = "&O Ñ¡Ïî"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_1
                    GetLocalizedString = "&S ÑùÊ½"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_2
                    GetLocalizedString = "&M ¸ü¶à"
                Case efnGUIStr_frmPrintGridFormatOptions_chkEnableAutoOrientation_Caption
                    GetLocalizedString = "Èç¹û±¨±í±ÈÖ½ÕÅ¿í£¬Ôò×Ô¶¯½«Ò³Ãæ·½Ïò¸ü¸ÄÎªË®Æ½¡£"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersSeparatorLine_Caption
                    GetLocalizedString = "±êÌâ·Ö¸ô·û"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintColumnsHeadersLines_Caption
                    GetLocalizedString = "ÁÐ±êÌâÐÐ"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersBorder_Caption
                    GetLocalizedString = "±êÌâ±ß½çÉ«"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintFixedColsBackground_Caption
                    GetLocalizedString = "¹Ì¶¨ÁÐ±³¾°É«"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersBackground_Caption
                    GetLocalizedString = "±êÌâ±³¾°É«"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintOtherBackgrounds_Caption
                    GetLocalizedString = "ÆäËû±³¾°É«"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintRowsLines_Caption
                    GetLocalizedString = "ÐÐÏß"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintColumnsDataLines_Caption
                    GetLocalizedString = "ÁÐÏß"
                Case efnGUIStr_frmPrintGridFormatOptions_chkPrintOuterBorder_Caption
                    GetLocalizedString = "Íâ²¿±ßÔµ"
                Case efnGUIStr_frmPrintGridFormatOptions_lblLineWidth_Caption
                    GetLocalizedString = "Ïß¿í:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblStyle_Caption
                    GetLocalizedString = "ÑùÊ½:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblOtherTextsFont_Caption
                    GetLocalizedString = "ÆäËûÎÄ±¾×ÖÌå:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblSubheadingFont_Caption
                    GetLocalizedString = "Ð¡±êÌâ×ÖÌå:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblHeadingFont_Caption
                    GetLocalizedString = "±êÌâ»ò±êÌâ×ÖÌå£º"
                Case efnGUIStr_frmPrintGridFormatOptions_lblPageNumbersFont_Caption
                    GetLocalizedString = "Ò³Âë×ÖÌå:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblPageNumbersFormat_Caption
                    GetLocalizedString = "Ò³Âë¸ñÊ½:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblPageNumbersPosition_Caption
                    GetLocalizedString = "Ò³ÂëÎ»ÖÃ:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblGridAlign_Caption
                    GetLocalizedString = "Íø¸ñ¶ÔÆë:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblColor_Caption
                    GetLocalizedString = "ÑÕÉ«:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblScalePercent_Caption
                    GetLocalizedString = "ÑùÊ½:"
                Case efnGUIStr_frmPrintGridFormatOptions_cboColor_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "´òÓ¡»úÄ¬ÈÏ"
                        Case 1
                            GetLocalizedString = "»Ò¶È"
                        Case 2
                            GetLocalizedString = "²ÊÉ«"
                    End Select
                Case efnGUIStr_frmPrintGridFormatOptions_cboPageNumbersPosition_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "ÓÒÏÂ½Ç"
                        Case 1
                            GetLocalizedString = "×óÏÂ½Ç"
                        Case 2
                            GetLocalizedString = "µ×²¿¾ÓÖÐ"
                        Case 3
                            GetLocalizedString = "ÓÒÉÏ½Ç"
                        Case 4
                            GetLocalizedString = "×óÉÏ½Ç"
                        Case 5
                            GetLocalizedString = "¶¥²¿¾ÓÖÐ"
                        Case 6
                            GetLocalizedString = "²»ÏÔÊ¾Ò³Âë"
                    End Select
                Case efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_List
                    Select Case nIndex
                        Case 0
                            GetLocalizedString = "¾ÓÖÐ"
                        Case 1
                            GetLocalizedString = "¾Ó×ó"
                        Case 2
                            GetLocalizedString = "¾ÓÓÒ"
                        Case 3
                            GetLocalizedString = "À­Éì"
                    End Select
                Case efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_Style
                    GetLocalizedString = "ÑùÊ½"
                Case efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_CustomStyle
                    GetLocalizedString = "×Ô¶¨ÒåÑùÊ½"
                Case efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_Customize
                    GetLocalizedString = "¶¨ÖÆ"
                Case efnGUIStr_frmPrintGridFormatOptions_lblSample_Caption
                    GetLocalizedString = "ÊÂÀý:"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Column
                    GetLocalizedString = "ÁÐ"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Element
                    GetLocalizedString = "ÔªËØ"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Data
                    GetLocalizedString = "Êý¾Ý"
                Case efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_ToolTipText
                    GetLocalizedString = "½öÔÚÊý¾ÝÍø¸ñ±ÈÒ³Ãæ¸üÕ­Ê±²ÅÓÐÐ§"
                Case efnGUIStr_frmPrintGridFormatOptions_txtLineWidth_ToolTipText
                    GetLocalizedString = "¸Ä±äÏßÌõ´ÖÏ¸£¨Ò»°ã£©"
                Case efnGUIStr_frmPrintGridFormatOptions_cmdHeadersBackgroundColor_ToolTipText
                    GetLocalizedString = "¸ü¸Ä±êÌâ£¨»ò¹Ì¶¨ÁÐ£©µÄ±³¾°ÑÕÉ«"
                Case efnGUIStr_frmPrintGridFormatOptions_txtLineWidthHeadersSeparatorLine_ToolTipText
                    GetLocalizedString = "Ò³Ã¼·Ö¸ôÏßºñ¶È"
                Case efnGUIStr_frmPrintGridFormatOptions_VariousChangeColorCommandButtons_ToolTipText
                    GetLocalizedString = "ÐÞ¸ÄÑÕÉ«"
                Case efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidth_Message
                    GetLocalizedString = "ÏßÌõµÄ¿í¶ÈÖµ±ØÐëÔÚ1µ½10Ö®¼ä"
                Case efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidthHeadersSeparatorLine_Message
                    GetLocalizedString = "ÏßÌõµÄ¿í¶ÈÖµ±ØÐëÔÚ1µ½20Ö®¼ä"
                ' frmSelectColumns
                Case efnGUIStr_frmSelectColumns_Caption
                    GetLocalizedString = "ÅäÖÃ¿É¼ûÁÐ"
                Case efnGUIStr_frmSelectColumns_lblTitle_Caption
                    GetLocalizedString = "&S Ñ¡ÔñÒªÏÔÊ¾µÄÁÐ:"
                Case efnGUIStr_frmSelectColumns_OneVisible_Message
                    GetLocalizedString = "±ØÐëÖÁÉÙÑ¡ÔñÒ»ÁÐ¡£"
                ' frmSettingGridDataProgress
                Case efnGUIStr_frmSettingGridDataProgress_lblMessage_Caption_Start
                    GetLocalizedString = "Éú³ÉÔ¤ÀÀ"
                Case efnGUIStr_frmSettingGridDataProgress_lblMessage_Caption_Progress
                    GetLocalizedString = "Éú³ÉÔ¤ÀÀ£¬Ò³"
                ' frmPrintPreview
                Case efnGUIStr_frmPrintPreview_Caption
                    GetLocalizedString = "´òÓ¡Ô¤ÀÀ"
                Case efnGUIStr_frmPrintPreview_mnuView2p_Caption
                    GetLocalizedString = "ÏÔÊ¾Á½Ò³"
                Case efnGUIStr_frmPrintPreview_mnuView3p_Caption
                    GetLocalizedString = "ÏÔÊ¾ÈýÒ³"
                Case efnGUIStr_frmPrintPreview_mnuView6p_Caption
                    GetLocalizedString = "ÏÔÊ¾ÁùÒ³"
                Case efnGUIStr_frmPrintPreview_mnuView12p_Caption
                    GetLocalizedString = "ÏÔÊ¾12Ò³"
                Case efnGUIStr_frmPrintPreview_mnuIconsAuto_Caption
                    GetLocalizedString = "×Ô¶¯"
                Case efnGUIStr_frmPrintPreview_mnuIconsSmall_Caption
                    GetLocalizedString = "Ð¡Í¼±ê"
                Case efnGUIStr_frmPrintPreview_mnuIconsMedium_Caption
                    GetLocalizedString = "ÖÐÍ¼±ê"
                Case efnGUIStr_frmPrintPreview_mnuIconsBig_Caption
                    GetLocalizedString = "µÈµÈÍ¼±ê"
                Case efnGUIStr_frmPrintPreview_lblPageOrientation_Caption
                    GetLocalizedString = "&P ÏÔÊ¾·½Ïò:"
                Case efnGUIStr_frmPrintPreview_lblView_Caption
                    GetLocalizedString = "&V ÏÔÊ¾:"
                Case efnGUIStr_frmPrintPreview_lblScalePercent_Caption
                    GetLocalizedString = "&S Ëõ·Å£º" 'Scale:"The first character must be a letter. . .
                Case efnGUIStr_frmPrintPreview_lblPage_Singular_Caption
                    GetLocalizedString = "Ò³:"
                Case efnGUIStr_frmPrintPreview_lblPage_Plural_Caption
                    GetLocalizedString = "×ÜÒ³Êý:"
                Case efnGUIStr_frmPrintPreview_lblPageCount_Caption
                    GetLocalizedString = "/"
                Case efnGUIStr_frmPrintPreview_lblPreparingDoc_Caption
                    GetLocalizedString = "ÕýÔÚÉú³É´òÓ¡Ô¤ÀÀ..."
                Case efnGUIStr_frmPrintPreview_cmdClose_Caption
                    GetLocalizedString = "&C ¹Ø±Õ"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Print
                    GetLocalizedString = "¿ªÊ¼´òÓ¡"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageSetup
                    GetLocalizedString = "Ò³ÃæÉèÖÃ"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Format
                    GetLocalizedString = "ÏÔÊ¾¸ñÊ½"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationPortrait
                    GetLocalizedString = "×ÝÏòÏÔÊ¾"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationLandscape
                    GetLocalizedString = "ºáÏòÏÔÊ¾"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewNormalSize
                    GetLocalizedString = "²é¿´Õý³£Ò³Ãæ´óÐ¡"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenWidth
                    GetLocalizedString = "Ò³Ãæµ÷ÕûÎªÆÁÄ»¿í¶È"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenHeight
                    GetLocalizedString = "Ò³Ãæµ÷Õûµ½ÆÁÄ»¸ß¶È"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewSeveralPages
                    GetLocalizedString = "ÏÔÊ¾¶àÒ³"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_DecreaseScale
                    GetLocalizedString = "¼õÉÙ×ÖÌåºÍÔªËØ´óÐ¡"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_IncreaseScale
                    GetLocalizedString = "Ôö¼Ó×ÖÌåºÍÔªËØ´óÐ¡"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_FirstPage
                    GetLocalizedString = "Ê×Ò³"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Singular
                    GetLocalizedString = "ÉÏÒ»Ò³"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Plural
                    GetLocalizedString = "Previous pages"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Singular
                    GetLocalizedString = "ÏÂÒ»Ò³"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Plural
                    GetLocalizedString = "Next pages"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_LastPage
                    GetLocalizedString = "×îºóÒ³"
                
                ' UserControls
                ' FontPicker
                Case efnGUIStr_FontPicker_ButtonToolTipTextDefault
                    GetLocalizedString = "Ñ¡Ôñ×ÖÌå"
                Case efnGUIStr_FontPicker_DrawSample_Bold
                    GetLocalizedString = "´ÖÌå"
                Case efnGUIStr_FontPicker_DrawSample_Italic
                    GetLocalizedString = "Ð±Ìå"
                ' FontSizeChanger
                Case efnGUIStr_FontSizeChanger_Extender_ToolTipText
                    GetLocalizedString = "µ¥»÷¼ÓºÅ»ò¼õºÅÐÞ¸Ä×ÖÌå´óÐ¡ £¬µ±Ç°×ÖÌå´óÐ¡£º"
                Case efnGUIStr_FontSizeChanger_mnuDefaultFontSize_Caption
                    GetLocalizedString = "ÉèÖÃÄ¬ÈÏÖµ"
                Case efnGUIStr_FontSizeChanger_btnMinus_ToolTipText
                    GetLocalizedString = "¼õÐ¡×ÖÌå´óÐ¡"
                Case efnGUIStr_FontSizeChanger_btnPlus_ToolTipText
                    GetLocalizedString = "Ôö¼Ó×ÖÌå´óÐ¡"
                ' FlexFn
                Case efnGUIStr_FlexFn_PrintButton_ToolTipText_Default
                    GetLocalizedString = "´òÓ¡"
                Case efnGUIStr_FlexFn_PrintPreviewButton_ToolTipText_Default
                    GetLocalizedString = "´òÓ¡ÉèÖÃºÍ´òÓ¡Ô¤ÀÀ"
                Case efnGUIStr_FlexFn_FindButton_ToolTipText_Default
                    GetLocalizedString = "²éÕÒ"
                Case efnGUIStr_FlexFn_CopyButton_ToolTipText_Default
                    GetLocalizedString = "¸´ÖÆ"
                Case efnGUIStr_FlexFn_SaveButton_ToolTipText_Default
                    GetLocalizedString = "±£´æµ½ÎÄ¼þ"
                Case efnGUIStr_FlexFn_GroupDataButton_ToolTipText_Default
                    GetLocalizedString = "ºÏ²¢ÏàÍ¬ÁÐ"
                Case efnGUIStr_FlexFn_GroupDataButtonPressed_ToolTipText_Default
                    GetLocalizedString = "²»ºÏ²¢ÏàÍ¬ÁÐ"
                Case efnGUIStr_FlexFn_ConfigColumnsButton_ToolTipText_Default
                    GetLocalizedString = "ÅäÖÃÒªÔÚ´Ë±¨±íÖÐÏÔÊ¾µÄÁÐ"
                Case efnGUIStr_FlexFn_ConfigColumnsButtonColsHidden_ToolTipText_Default
                    GetLocalizedString = "ÅäÖÃÁÐ£¨Òþ²ØÁÐ£©"
                Case efnGUIStr_FlexFn_CopyCellMenuCaption_Default
                    GetLocalizedString = "µ¥Ôª¸ñ"
                Case efnGUIStr_FlexFn_CopyRowMenuCaption_Default
                    GetLocalizedString = "ÐÐ"
                Case efnGUIStr_FlexFn_CopyColumnMenuCaption_Default
                    GetLocalizedString = "ÁÐ"
                Case efnGUIStr_FlexFn_CopyAllMenuCaption_Default
                    GetLocalizedString = "È«²¿"
                Case efnGUIStr_FlexFn_mnuCopyParent_Caption
                    GetLocalizedString = "¸´ÖÆ..."
                ' History
                Case efnGUIStr_History_mnuDelete_Caption1
                    GetLocalizedString = "É¾³ý"
                Case efnGUIStr_History_mnuDelete_Caption2
                    GetLocalizedString = "´ÓÀúÊ·¼ÍÂ¼"
                Case efnGUIStr_History_ToolTipTextStart_Default
                    GetLocalizedString = "×ªµ½"
                Case efnGUIStr_History_ToolTipTextSelect_Default
                    GetLocalizedString = "(»òÕßµ¥»÷ÓÒ¼üÑ¡Ôñ)"
                Case efnGUIStr_History_BackButtonToolTipText_Default
                    GetLocalizedString = "×ªµ½ÉÏÒ»Ìõ(»òÕßµ¥»÷ÓÒ¼üÑ¡Ôñ)"
                Case efnGUIStr_History_ForwardButtonToolTipText_Default
                    GetLocalizedString = "×ªµ½ÏÂÒ»Ìõ(»òÕßµ¥»÷ÓÒ¼üÑ¡Ôñ)"
                ' DateEnter
                Case efnGUIStr_DateEnter_ToolTipTextStart_Default
                    GetLocalizedString = "°´¸ñÊ½ÊäÈëÈÕÆÚ"
                Case efnGUIStr_DateEnter_ToolTipTextEnd_Default
                    GetLocalizedString = "»òµ¥»÷¼ýÍ·Ñ¡Ôñ"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxTitle
                    GetLocalizedString = "ÈÕÆÚÊäÈë´íÎó"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError1
                    GetLocalizedString = "ÄúÃ»ÓÐÔÚÈÕÆÚÌõÄ¿ÖÐÊäÈëÈÕÆÚ¡£"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError2
                    GetLocalizedString = "ÌìÊý²»ÄÜÎªÁã¡£"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError3
                    GetLocalizedString = "²»ÄÜ³¬¹ý31Ìì"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError4
                    GetLocalizedString = "ÄúÃ»ÓÐÔÚÈÕÆÚÌõÄ¿ÖÐÊäÈëÔÂ·Ý¡£"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError5
                    GetLocalizedString = "ÔÂ·ÝµÄÖµ±ØÐëÔÚ1µ½12Ö®¼ä¡£"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError6
                    GetLocalizedString = "ÄúÃ»ÓÐÔÚÈÕÆÚÌõÄ¿ÖÐÊäÈëÄê·Ý¡£"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError7
                    GetLocalizedString = "ÈÕÆÚ²»ÄÜµÍÓÚ"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError8
                    GetLocalizedString = "ÈÕÆÚ²»ÄÜ´óÓÚ"
                
                ' Class modules
                ' cGridHandler
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString1
                    GetLocalizedString = "°´´ËÁÐÅÅÐò" 'Order by this column"
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString2
                    GetLocalizedString = "Order by"
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString3
                    GetLocalizedString = "ÕýÏòÅÅÐò" 'ascending"
                Case efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString4
                    GetLocalizedString = "·´ÏòÅÅÐò" 'descending"
            
                ' FlexFnObject
                Case efnGUIStr_FlexFnObject_FindTextInGrid_InputBoxEnterTextMessageTitle
                    GetLocalizedString = "²éÕÒÎÄ×Ö"
                Case efnGUIStr_FlexFnObject_FindTextInGrid_InputBoxEnterTextMessage
                    GetLocalizedString = "ÇëÊäÈëÒª²éÕÒµÄÎÄ×Ö:"
                Case efnGUIStr_FlexFnObject_FindTextInGrid_MsgboxTextNotFound
                    GetLocalizedString = "Ã»ÓÐÕÒµ½Òª²éÕÒµÄÎÄ×Ö."
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1a
                    GetLocalizedString = "ÎÄ¼þ"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1b
                    GetLocalizedString = "ÒÑ¾­´æÔÚ£¬ÊÇ·ñ¸²¸Ç£¿"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2a
                    GetLocalizedString = "ÎÄ¼þ²»ÄÜ±»Ìæ»»£¬Ëü¿ÉÒÔÓÃExcel´ò¿ª¡£ÒªÓÃÏàÍ¬µÄÃû³Æ±£´æËü£¬±ØÐëÏÈ¹Ø±ÕËü¡£"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2b
                    GetLocalizedString = "ÊÇ·ñÖØÊÔ£¿£¨ÔÚ¹Ø±Õºó°´¡°ÊÇ¡±£¬¡°·ñ¡±Ñ¡ÔñÁíÒ»¸öÃû³Æ£¬»ò¡°È¡Ïû¡±È¡Ïû²Ù×÷"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox3
                    GetLocalizedString = "ÊÇ·ñÏÖÔÚ´ò¿ª±£´æºóµÄÎÄ¼þ£¿"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox4
                    GetLocalizedString = "ÎÄ¼þ±£´æÔÚ"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox5
                    GetLocalizedString = "·¢Éú´íÎó:"
                Case efnGUIStr_FlexFnObject_PageNumbersFormatStrings_Count
                    GetLocalizedString = "15"
                Case efnGUIStr_FlexFnObject_PageNumbersFormatStrings
                    Select Case nIndex
                        Case 1 ' the index starts in 1, the 0 is retrieved from efnGUIStr_General_PageNumbersFormatString_Default
                            GetLocalizedString = "Ò³ #"
                        Case 2
                            GetLocalizedString = "Ò³ #"
                        Case 3
                            GetLocalizedString = "Ò³ # / N"
                        Case 4
                            GetLocalizedString = "Ò³ # / N"
                        Case 5
                            GetLocalizedString = "Ò³. #"
                        Case 6
                            GetLocalizedString = "# / N"
                        Case 7
                            GetLocalizedString = "Ò³. # / N"
                        Case 8
                            GetLocalizedString = "#/N"
                        Case 9
                            GetLocalizedString = "Ò³. #/N"
                        Case 10
                            GetLocalizedString = "Ò³ #/N"
                        Case 11
                            GetLocalizedString = "Ò³ #/N"
                        Case 12
                            GetLocalizedString = "# / N"
                        Case 13
                            GetLocalizedString = "Ò³. # / N"
                        Case 14
                            GetLocalizedString = "Ò³ # / N"
                        Case 15
                            GetLocalizedString = "Ò³ # / N"
                    End Select
                ' cPrinterEx
                Case efnGUIStr_cPrinterEx_PrintDocumentTooLongWarning_MsgBoxWarning
                    GetLocalizedString = "ÎÄµµÌ«³¤£¬½«ÎÞ·¨ÍêÈ«´òÓ¡¡£"
                ' Bas modules
                ' mGlobals
                Case efnGUIStr_mGlobals_ValidFileName_DefaultFileName
                    GetLocalizedString = "Î´ÃüÃû"
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
                    GetLocalizedString = "Se copió el texto"
                ' frmConfigHistory
                Case efnGUIStr_frmConfigHistory_Caption
                    GetLocalizedString = "Configuración de historial"
                Case efnGUIStr_frmConfigHistory_chkRememberHistory_Caption
                    GetLocalizedString = "Recordar el historial a través de sesiones"
                Case efnGUIStr_frmConfigHistory_HelpMessageTitle
                    GetLocalizedString = "Historial"
                Case efnGUIStr_frmConfigHistory_HelpMessage
                    GetLocalizedString = "Indica si el programa recordará lo buscado o seleccionado la próxima vez que lo corra, luego de cerrarlo." & vbCrLf & vbCrLf & "Puede eliminar el historial actual haciendo click en uno de los botones." & vbCrLf & "También se puede eliminar un solo elemento de un historial haciendo click con el botón derecho del mouse sobre el mismo (en la lista, antes de entrar a esta pantalla de configuración)."
                Case efnGUIStr_frmConfigHistory_cmdEraseContext_Caption
                    GetLocalizedString = "Eliminar el historial para este contexto"
                Case efnGUIStr_frmConfigHistory_cmdEraseAll_Caption
                    GetLocalizedString = "Eliminar todos los historiales"
                ' frmCopyGridTextOptions
                Case efnGUIStr_frmCopyGridTextOptions_Caption
                    GetLocalizedString = "Opciones de copia de texto"
                Case efnGUIStr_frmCopyGridTextOptions_lblColumnsSeparationMode_Caption
                    GetLocalizedString = "Separación de las columnas:"
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
                    GetLocalizedString = "Seleccionar qué columnas incluir:"
                Case efnGUIStr_EnterColumnSeparatorMessage
                    GetLocalizedString = "Por favor ingrese el texto o caracter separador de columnas"
                Case efnGUIStr_EnterColumnSeparatorMessageTitle
                    GetLocalizedString = "Ingresar separador"
                Case efnGUIStr_SelectFontMessage
                    GetLocalizedString = "Necesita seleccionar qué fuente va a usar en el programa destino donde lo va a pegar." & vbCrLf & "El alineado de las columnas será aproximado."
                ' frmPrintGridFormatOptions
                Case efnGUIStr_frmPrintGridFormatOptions_Caption
                    GetLocalizedString = "Formato de impresión"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_0
                    GetLocalizedString = "Opciones"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_1
                    GetLocalizedString = "Estilo"
                Case efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_2
                    GetLocalizedString = "Otras"
                Case efnGUIStr_frmPrintGridFormatOptions_chkEnableAutoOrientation_Caption
                    GetLocalizedString = "Cambiar automáticamente la orientación de la página a horizontal si el listado es más ancho que el papel."
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
                    GetLocalizedString = "Grosor de líneas:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblStyle_Caption
                    GetLocalizedString = "Estilo:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblOtherTextsFont_Caption
                    GetLocalizedString = "Fuente otros textos:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblSubheadingFont_Caption
                    GetLocalizedString = "Fuente del sub-encabezado:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblHeadingFont_Caption
                    GetLocalizedString = "Fuente del encabezado o título:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblPageNumbersFont_Caption
                    GetLocalizedString = "Fuente de números de página:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblPageNumbersFormat_Caption
                    GetLocalizedString = "Formato de números de página:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblPageNumbersPosition_Caption
                    GetLocalizedString = "Posición de números de página:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblGridAlign_Caption
                    GetLocalizedString = "Alineación de grilla de datos:"
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
                Case efnGUIStr_frmPrintGridFormatOptions_cboPageNumbersPosition_List
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
                            GetLocalizedString = "No colocar números de página"
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
                    GetLocalizedString = "Personalizado"
                Case efnGUIStr_frmPrintGridFormatOptions_lblSample_Caption
                    GetLocalizedString = "Ejemplo:"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Column
                    GetLocalizedString = "Columna"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Element
                    GetLocalizedString = "Elemento"
                Case efnGUIStr_frmPrintGridFormatOptions_DrawSample_Data
                    GetLocalizedString = "Dato"
                Case efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_ToolTipText
                    GetLocalizedString = "Sólo tiene efecto cuando la grilla de datos es más angosta que la página"
                Case efnGUIStr_frmPrintGridFormatOptions_txtLineWidth_ToolTipText
                    GetLocalizedString = "Cambiar grosor de líneas (general)"
                Case efnGUIStr_frmPrintGridFormatOptions_cmdHeadersBackgroundColor_ToolTipText
                    GetLocalizedString = "Cambiar color de fondo de encabezados (y/o columnas fijas)"
                Case efnGUIStr_frmPrintGridFormatOptions_txtLineWidthHeadersSeparatorLine_ToolTipText
                    GetLocalizedString = "Grosor de línea separadora de encabezados"
                Case efnGUIStr_frmPrintGridFormatOptions_VariousChangeColorCommandButtons_ToolTipText
                    GetLocalizedString = "Cambiar color"
                Case efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidth_Message
                    GetLocalizedString = "El valor del grosor de las líneas debe estar entre 1 y 10"
                Case efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidthHeadersSeparatorLine_Message
                    GetLocalizedString = "El valor del grosor de las líneas debe estar entre 1 y 20"
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
                    GetLocalizedString = "Generando vista previa, página"
                ' frmPrintPreview
                Case efnGUIStr_frmPrintPreview_Caption
                    GetLocalizedString = "Vista preliminar de impresión"
                Case efnGUIStr_frmPrintPreview_mnuView2p_Caption
                    GetLocalizedString = "Ver 2 páginas"
                Case efnGUIStr_frmPrintPreview_mnuView3p_Caption
                    GetLocalizedString = "Ver 3 páginas"
                Case efnGUIStr_frmPrintPreview_mnuView6p_Caption
                    GetLocalizedString = "Ver 6 páginas"
                Case efnGUIStr_frmPrintPreview_mnuView12p_Caption
                    GetLocalizedString = "Ver 12 páginas"
                Case efnGUIStr_frmPrintPreview_mnuIconsAuto_Caption
                    GetLocalizedString = "Automático"
                Case efnGUIStr_frmPrintPreview_mnuIconsSmall_Caption
                    GetLocalizedString = "Íconos pequeños"
                Case efnGUIStr_frmPrintPreview_mnuIconsMedium_Caption
                    GetLocalizedString = "Íconos  medianos"
                Case efnGUIStr_frmPrintPreview_mnuIconsBig_Caption
                    GetLocalizedString = "Íconos grandes"
                Case efnGUIStr_frmPrintPreview_lblPageOrientation_Caption
                    GetLocalizedString = "Orientación de página:"
                Case efnGUIStr_frmPrintPreview_lblView_Caption
                    GetLocalizedString = "Ver:"
                Case efnGUIStr_frmPrintPreview_lblScalePercent_Caption
                    GetLocalizedString = "Escala:"
                Case efnGUIStr_frmPrintPreview_lblPage_Singular_Caption
                    GetLocalizedString = "Página:"
                Case efnGUIStr_frmPrintPreview_lblPage_Plural_Caption
                    GetLocalizedString = "Páginas:"
                Case efnGUIStr_frmPrintPreview_lblPageCount_Caption
                    GetLocalizedString = "de"
                Case efnGUIStr_frmPrintPreview_lblPreparingDoc_Caption
                    GetLocalizedString = "Generando vista preliminar..."
                Case efnGUIStr_frmPrintPreview_cmdClose_Caption
                    GetLocalizedString = "Cerrar vista preliminar"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Print
                    GetLocalizedString = "Imprimir"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageSetup
                    GetLocalizedString = "Configurar página"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Format
                    GetLocalizedString = "Formato"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationPortrait
                    GetLocalizedString = "Orientación vertical"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationLandscape
                    GetLocalizedString = "Orientación horizontal"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewNormalSize
                    GetLocalizedString = "Ver tamaño de página normal"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenWidth
                    GetLocalizedString = "Ver página ajustada al ancho de la pantalla"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenHeight
                    GetLocalizedString = "Ver página ajustada al alto de la pantalla"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewSeveralPages
                    GetLocalizedString = "Ver varias páginas a la vez"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_DecreaseScale
                    GetLocalizedString = "Disminuir tamaño de fuentes y elementos"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_IncreaseScale
                    GetLocalizedString = "Aumentar tamaño de fuentes y elementos"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_FirstPage
                    GetLocalizedString = "Primera página"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Singular
                    GetLocalizedString = "Página anterior"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Plural
                    GetLocalizedString = "Páginas anteriores"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Singular
                    GetLocalizedString = "Página siguiente"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Plural
                    GetLocalizedString = "Próximas páginas"
                Case efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_LastPage
                    GetLocalizedString = "Última página"
                
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
                    GetLocalizedString = "Haga clic en los signos + y - si desea cambiar el tamaño de la letra (el tamaño actual es "
                Case efnGUIStr_FontSizeChanger_mnuDefaultFontSize_Caption
                    GetLocalizedString = "Poner valor por defecto"
                Case efnGUIStr_FontSizeChanger_btnMinus_ToolTipText
                    GetLocalizedString = "Disminuir el tamaño de la letra"
                Case efnGUIStr_FontSizeChanger_btnPlus_ToolTipText
                    GetLocalizedString = "Aumentar el tamaño de la letra"
                ' FlexFn
                Case efnGUIStr_FlexFn_PrintButton_ToolTipText_Default
                    GetLocalizedString = "Imprimir"
                Case efnGUIStr_FlexFn_PrintPreviewButton_ToolTipText_Default
                    GetLocalizedString = "Configuración de impresión y vista preliminar"
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
                    GetLocalizedString = "Configurar qué columnas mostrar en este listado"
                Case efnGUIStr_FlexFn_ConfigColumnsButtonColsHidden_ToolTipText_Default
                    GetLocalizedString = "Configurar columnas (hay columnas ocultas)"
                Case efnGUIStr_FlexFn_CopyCellMenuCaption_Default
                    GetLocalizedString = "Celda"
                Case efnGUIStr_FlexFn_CopyRowMenuCaption_Default
                    GetLocalizedString = "Línea"
                Case efnGUIStr_FlexFn_CopyColumnMenuCaption_Default
                    GetLocalizedString = "Columna"
                Case efnGUIStr_FlexFn_CopyAllMenuCaption_Default
                    GetLocalizedString = "Todo"
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
                    GetLocalizedString = "(o clic con el botón derecho para seleccionar)"
                Case efnGUIStr_History_BackButtonToolTipText_Default
                    GetLocalizedString = "Ir a ítem anterior (o click con el botón derecho para seleccionar)"
                Case efnGUIStr_History_ForwardButtonToolTipText_Default
                    GetLocalizedString = "Ir a ítem siguiente (o click con el botón derecho para seleccionar)"
                ' DateEnter
                Case efnGUIStr_DateEnter_ToolTipTextStart_Default
                    GetLocalizedString = "Ingrese la fecha en formato"
                Case efnGUIStr_DateEnter_ToolTipTextEnd_Default
                    GetLocalizedString = "o haga clic en la flecha para seleccionar"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxTitle
                    GetLocalizedString = "Error en ingreso de fecha"
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError1
                    GetLocalizedString = "No ingresó el día en el ingreso de fecha."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError2
                    GetLocalizedString = "El día no puede ser cero."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError3
                    GetLocalizedString = "El día no puede ser mayor a 31."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError4
                    GetLocalizedString = "No ingresó el mes en el ingreso de fecha."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError5
                    GetLocalizedString = "El valor del mes debe estar entre 1 y 12."
                Case efnGUIStr_DateEnter_Validate1_MsgBoxError6
                    GetLocalizedString = "No ingresó el año en el ingreso de fecha."
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
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1a
                    GetLocalizedString = "El archivo"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1b
                    GetLocalizedString = "ya existe, ¿desea sobreescribirlo?"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2a
                    GetLocalizedString = "El archivo no se puede reemplazar, es posible que esté abierto con Excel, para guardarlo con el mismo nombre debe cerrarlo antes."
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2b
                    GetLocalizedString = "¿Reintentar? (Presione 'Sí' luego de cerrarlo, 'No' para elegir otro nombre, o 'Cancelar' para cancelar la operación."
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox3
                    GetLocalizedString = "¿Desea abrir ahora con Excel el archivo guardado?"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox4
                    GetLocalizedString = "El archivo se guardó en"
                Case efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox5
                    GetLocalizedString = "Se produjo un error:"
                Case efnGUIStr_FlexFnObject_PageNumbersFormatStrings_Count
                    GetLocalizedString = "15"
                Case efnGUIStr_FlexFnObject_PageNumbersFormatStrings
                    Select Case nIndex
                        Case 1 ' the index starts in 1, the 0 is retrieved from efnGUIStr_General_PageNumbersFormatString_Default
                            GetLocalizedString = "Página #"
                        Case 2
                            GetLocalizedString = "página #"
                        Case 3
                            GetLocalizedString = "Página # de N"
                        Case 4
                            GetLocalizedString = "página # de N"
                        Case 5
                            GetLocalizedString = "Pág. #"
                        Case 6
                            GetLocalizedString = "# de N"
                        Case 7
                            GetLocalizedString = "Pág. # de N"
                        Case 8
                            GetLocalizedString = "#/N"
                        Case 9
                            GetLocalizedString = "Pág. #/N"
                        Case 10
                            GetLocalizedString = "Página #/N"
                        Case 11
                            GetLocalizedString = "página #/N"
                        Case 12
                            GetLocalizedString = "# / N"
                        Case 13
                            GetLocalizedString = "Pág. # / N"
                        Case 14
                            GetLocalizedString = "Página # / N"
                        Case 15
                            GetLocalizedString = "página # / N"
                    End Select
                ' cPrinterEx
                Case efnGUIStr_cPrinterEx_PrintDocumentTooLongWarning_MsgBoxWarning
                    GetLocalizedString = "Documento demasiado extenso, no se imprimirá completo."
                ' Bas modules
                ' mGlobals
                Case efnGUIStr_mGlobals_ValidFileName_DefaultFileName
                    GetLocalizedString = "Sin título"
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
                Case efnGUIStr_frmPrintGridFormatOptions_lblPageNumbersFont_Caption
                    GetLocalizedString = "Page numbers font:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblPageNumbersFormat_Caption
                    GetLocalizedString = "Page numbers format:"
                Case efnGUIStr_frmPrintGridFormatOptions_lblPageNumbersPosition_Caption
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
                Case efnGUIStr_frmPrintGridFormatOptions_cboPageNumbersPosition_List
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
                    GetLocalizedString = "Customized"
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
                    GetLocalizedString = "The thickness value of the lines must be between 1 and 10"
                Case efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidthHeadersSeparatorLine_Message
                    GetLocalizedString = "The thickness value of the lines must be between 1 and 20"
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
                    GetLocalizedString = "Ver 12 páginas"
                Case efnGUIStr_frmPrintPreview_mnuIconsAuto_Caption
                    GetLocalizedString = "Automatic"
                Case efnGUIStr_frmPrintPreview_mnuIconsSmall_Caption
                    GetLocalizedString = "Small icons"
                Case efnGUIStr_frmPrintPreview_mnuIconsMedium_Caption
                    GetLocalizedString = "Medium icons"
                Case efnGUIStr_frmPrintPreview_mnuIconsBig_Caption
                    GetLocalizedString = "Large icons"
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
                Case efnGUIStr_frmPrintPreview_lblPreparingDoc_Caption
                    GetLocalizedString = "Generating print preview..."
                Case efnGUIStr_frmPrintPreview_cmdClose_Caption
                    GetLocalizedString = "Close print preview"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Print
                    GetLocalizedString = "Print"
                Case efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageSetup
                    GetLocalizedString = "Page setup"
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



