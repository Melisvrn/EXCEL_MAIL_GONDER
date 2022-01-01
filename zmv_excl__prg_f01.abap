*&---------------------------------------------------------------------*
*&  Include           ZMV_EXCL__PRG_F01
*&---------------------------------------------------------------------*

FORM GET_DATA .

  SELECT MARA~MATNR MARA~ERSDA MARA~ERNAM
         MARA~LAEDA MARD~WERKS MARD~LGORT MARC~EKGRP MARC~AUSME
             FROM MARA INNER JOIN MARC ON MARA~MATNR EQ MARC~MATNR
             INNER JOIN MARD ON MARC~MATNR EQ MARD~MATNR
             AND MARC~WERKS EQ MARD~WERKS
             INTO CORRESPONDING FIELDS OF TABLE GT_ITAB
             WHERE MARA~MATNR IN S_MATNR
             AND MARD~WERKS EQ P_WERKS .

  SELECT SINGLE * FROM MARD
    WHERE WERKS EQ P_WERKS.
  IF  P_WERKS IS INITIAL.
    MESSAGE ' UYARI !!  Üretim Yeri Girilmeden Malzeme Görüntülenemez ' TYPE 'I'.
    STOP.
  ENDIF.

ENDFORM.
*&---------------------------------------------------------------------*
*&      Form  CREATE_FİELDCAT_0100
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM CREATE_FIELDCAT_0100 .

  CALL FUNCTION 'LVC_FIELDCATALOG_MERGE'
    EXPORTING
*     I_BUFFER_ACTIVE        =
      I_STRUCTURE_NAME       = 'ZMV_EXCL_STR'
    CHANGING
      CT_FIELDCAT            = FIELDCAT
    EXCEPTIONS
      INCONSISTENT_INTERFACE = 1
      PROGRAM_ERROR          = 2
      OTHERS                 = 3.

  LOOP AT FIELDCAT INTO GS_FIELDCAT.
    CLEAR GS_FIELDCAT.
    CASE   GS_FIELDCAT-FIELDNAME.
      WHEN 'AUSME' .
        GS_FIELDCAT-EDIT = 'X'.
      WHEN OTHERS.
    ENDCASE.
    MODIFY FIELDCAT FROM GS_FIELDCAT.
  ENDLOOP.



ENDFORM.
*&---------------------------------------------------------------------*
*&      Form  DİSPLAY_0100
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM DISPLAY_0100 .

  CREATE OBJECT ALV_0100
    EXPORTING
      I_PARENT = ALV_GRID_0100.

  GS_LAYOUT_0100-CWIDTH_OPT = 'X'. "UZUNLUK
  GS_LAYOUT_0100-INFO_FNAME  = 'CLINE' . "RENK
  GS_LAYOUT_0100-ZEBRA = 'X'. "ARKA PLAN

  CALL METHOD ALV_0100->SET_TABLE_FOR_FIRST_DISPLAY
    EXPORTING
      IS_LAYOUT          = GS_LAYOUT_0100
      I_BUFFER_ACTIVE    = SPACE
      I_BYPASSING_BUFFER = 'X'
      I_SAVE             = 'A'
    CHANGING
      IT_OUTTAB          = GT_ITAB[]
      IT_FIELDCATALOG    = FIELDCAT.
ENDFORM.

*&      Form  EXCEL_FORMAT_DOWNLOAD
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM EXCEL_FORMAT_DOWNLOAD .

  CREATE OBJECT EXCEL
    'EXCEL.APPLICATION'.
     "Yaratılacak nesnenin excel formatında olmasını sağlar.

  CALL METHOD OF
    EXCEL
      'WORKBOOKS' = WORKBOOK.

  SET PROPERTY OF
      EXCEL
       'VISIBLE' = 0. " Excel dosyasını arka plana koy.

  SET PROPERTY OF
       EXCEL
       'VISIBLE' = 1. " Excel dosyasını ön plana koy.

  CALL METHOD OF
    WORKBOOK
    'add'.

  CALL METHOD OF
      EXCEL
      'Worksheets' = SHEET
    EXPORTING
      #1           = 1.

  CALL METHOD OF
    SHEET
    'Activate'.

  SET PROPERTY OF SHEET 'Name' = 'SAYFA1'.
  "BAŞLIK SATIRLARI
  CALL METHOD OF EXCEL 'RANGE' = CELL EXPORTING #1 = 'A1'.
  SET PROPERTY OF CELL 'VALUE' = 'MALZEME NUMARASI'.

  CALL METHOD OF EXCEL 'RANGE' = CELL EXPORTING #1 = 'B1'.
  SET PROPERTY OF CELL 'VALUE' = 'YARATMA TARİHİ'.

  CALL METHOD OF EXCEL 'RANGE' = CELL EXPORTING #1 = 'C1'.
  SET PROPERTY OF CELL 'VALUE' = 'NESNEYİ YARATAN SORUMLUNUN ADI'.

  CALL METHOD OF EXCEL 'RANGE' = CELL EXPORTING #1 = 'D1'.
  SET PROPERTY OF CELL 'VALUE' = 'SON DEĞİŞİKLİK TARİHİ'..

  CALL METHOD OF EXCEL 'RANGE' = CELL EXPORTING #1 = 'E1'.
  SET PROPERTY OF CELL 'VALUE' = 'ÜRETİM YERİ'.

  CALL METHOD OF EXCEL 'RANGE' = CELL EXPORTING #1 = 'F1'.
  SET PROPERTY OF CELL 'VALUE' = 'DEPO YERİ'.

  CALL METHOD OF EXCEL 'RANGE' = CELL EXPORTING #1 = 'G1'.
  SET PROPERTY OF CELL 'VALUE' = 'SATIN ALMA GRUBU'.

  CALL METHOD OF EXCEL 'RANGE' = CELL EXPORTING #1 = 'H1'.
  SET PROPERTY OF CELL 'VALUE' = 'ÇIKIŞ ÖLÇÜ BİRİMİ'.

  DATA LV_INDEX TYPE STRING VALUE 1 .
  DATA LV_ST TYPE STRING.

  LOOP AT GT_ITAB INTO GS_ITAB.

    LV_INDEX = LV_INDEX + 1.

    CONCATENATE 'A' LV_INDEX INTO LV_ST.

    CALL METHOD OF EXCEL 'ROWS' = ROW EXPORTING #1 = LV_INDEX .
    CALL METHOD OF ROW 'INSERT' NO FLUSH.
    " No flush : sonraki komut bir OLE ifadesi olmasa bile toplama işlemine devam eder.
    CALL METHOD OF EXCEL 'RANGE' = CELL NO FLUSH EXPORTING #1 = LV_ST .
    SET PROPERTY OF CELL 'VALUE' = GS_ITAB-MATNR NO FLUSH.

    CLEAR LV_ST.
    CONCATENATE 'B' LV_INDEX INTO LV_ST.

    CALL METHOD OF EXCEL 'RANGE' = CELL NO FLUSH EXPORTING #1 = LV_ST .
    SET PROPERTY OF CELL 'VALUE' = GS_ITAB-ERSDA NO FLUSH.

    CLEAR LV_ST.
    CONCATENATE 'C' LV_INDEX INTO LV_ST.

    CALL METHOD OF EXCEL 'RANGE' = CELL NO FLUSH EXPORTING #1 = LV_ST .
    SET PROPERTY OF CELL 'VALUE' = GS_ITAB-ERNAM NO FLUSH.

    CLEAR LV_ST.
    CONCATENATE 'D' LV_INDEX INTO LV_ST.

    CALL METHOD OF EXCEL 'RANGE' = CELL NO FLUSH EXPORTING #1 = LV_ST .
    SET PROPERTY OF CELL 'VALUE' = GS_ITAB-LAEDA NO FLUSH.

    CLEAR LV_ST.
    CONCATENATE 'E' LV_INDEX INTO LV_ST.

    CALL METHOD OF EXCEL 'RANGE' = CELL NO FLUSH EXPORTING #1 = LV_ST .
    SET PROPERTY OF CELL 'VALUE' = GS_ITAB-WERKS NO FLUSH.

    CLEAR LV_ST.
    CONCATENATE 'F' LV_INDEX INTO LV_ST.

    CALL METHOD OF EXCEL 'RANGE' = CELL NO FLUSH EXPORTING #1 = LV_ST .
    SET PROPERTY OF CELL 'VALUE' = GS_ITAB-LGORT NO FLUSH.


    CLEAR LV_ST.
    CONCATENATE 'G' LV_INDEX INTO LV_ST.

    CALL METHOD OF EXCEL 'RANGE' = CELL NO FLUSH EXPORTING #1 = LV_ST .
    SET PROPERTY OF CELL 'VALUE' = GS_ITAB-EKGRP NO FLUSH.

    CLEAR LV_ST.
    CONCATENATE 'H' LV_INDEX INTO LV_ST.

    CALL METHOD OF EXCEL 'RANGE' = CELL NO FLUSH EXPORTING #1 = LV_ST.
    SET PROPERTY OF CELL 'VALUE' = GS_ITAB-AUSME NO FLUSH.

    CLEAR GS_ITAB.

  ENDLOOP.

*  * Tüm nesneleri serbest bırakın
  FREE OBJECT CELL.
  FREE OBJECT WORKBOOK.
  FREE OBJECT EXCEL.
  EXCEL-HANDLE = -1.
  FREE OBJECT ROW.

ENDFORM.
*&---------------------------------------------------------------------*
*&      Form  EXCEL_MAIL_GONDER
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM EXCEL_MAIL_GONDER.

  TRY. 
      SEND_REQUEST = CL_BCS=>CREATE_PERSISTENT( ).
      APPEND 'MERHABA. EXCEL DOSYAM EKTEDİR. İYİ ÇALIŞMALAR DİLERİM' TO TEXT.
      DOCUMENT = CL_DOCUMENT_BCS=>CREATE_DOCUMENT( "e-posta gövdesini ek olarak oluşturur.
      I_TYPE = 'RAW'
      I_TEXT = TEXT
      I_LENGTH = '12'
      I_SUBJECT = 'Test Immediate' ).
      DATA: LV_STR_LINE LIKE SOLIX-LINE.


*      GS_ITAB-COL1 = 'First'.
*      GS_ITAB-COL2 = 'Middle'.
*      GS_ITAB-COL3 = 'Last'.
*      APPEND GS_ITAB TO GT_ITAB.

      CONCATENATE LV_STRING
      'MALZEME NUMARASI' LC_TABDEL
      'YARATMA TARİHİ' LC_TABDEL
      'NESNEYİ YARATAN SORUMLUNUN ADI' LC_TABDEL
      'SON DEĞİŞİKLİK TARİHİ' LC_TABDEL
      'ÜRETİM YERİ' LC_TABDEL
      'DEPO YERİ' LC_TABDEL
      'SATIN ALMA GRUBU' LC_TABDEL
      'ÇIKIŞ ÖLÇÜ BİRİMİ'  LC_NEWLINE INTO LV_STRING2.

      LOOP AT GT_ITAB INTO GS_ITAB.

        CONCATENATE GS_ITAB-MATNR GS_ITAB-ERSDA GS_ITAB-ERNAM GS_ITAB-LAEDA
       GS_ITAB-WERKS GS_ITAB-LGORT GS_ITAB-EKGRP GS_ITAB-AUSME INTO LV_STRING SEPARATED BY
       LC_TABDEL.

        IF SY-TABIX EQ 1.
          CONCATENATE LV_STRING2 LV_STRING INTO LV_STRING2.
        ELSE.
          CONCATENATE LV_STRING2 LV_STRING INTO LV_STRING2 SEPARATED BY LC_NEWLINE.
        ENDIF.
        CLEAR: GS_ITAB.
      ENDLOOP.
      TRY.
          CL_BCS_CONVERT=>STRING_TO_SOLIX(
          EXPORTING
            IV_STRING = LV_STRING2
            IV_CODEPAGE = '4103' 
            IV_ADD_BOM = 'X' 
            IMPORTING
              ET_SOLIX = BINARY_CONTENT
              EV_SIZE = SIZE ).
        CATCH CX_BCS.
          MESSAGE E445(SO) INTO LV_MESSAGE .

      ENDTRY.

      CALL METHOD DOCUMENT->ADD_ATTACHMENT
        EXPORTING
          I_ATTACHMENT_TYPE    = 'XLS'  
          I_ATTACHMENT_SUBJECT = 'EXCEL DOSYASI'    
          I_ATT_CONTENT_HEX    = BINARY_CONTENT. 
      CALL METHOD SEND_REQUEST->SET_DOCUMENT( DOCUMENT ).
      SENDER = CL_SAPUSER_BCS=>CREATE( SY-UNAME ).

      CALL METHOD SEND_REQUEST->SET_SENDER
        EXPORTING
          I_SENDER = SENDER.
      RECIPIENT = CL_CAM_ADDRESS_BCS=>CREATE_INTERNET_ADDRESS('tugba.bulut@vbap.com.tr' ).

      CALL METHOD SEND_REQUEST->ADD_RECIPIENT
        EXPORTING
          I_RECIPIENT = RECIPIENT
          I_EXPRESS   = 'X'.
      SEND_REQUEST->SET_SEND_IMMEDIATELY( 'X' ).

      CALL METHOD SEND_REQUEST->SEND(
        EXPORTING
          I_WITH_ERROR_SCREEN = 'X'
        RECEIVING
          RESULT              = SENT_TO_ALL ).
      IF SENT_TO_ALL = 'X'.
        WRITE TEXT-003.
      ENDIF.

      COMMIT WORK.

    CATCH CX_BCS INTO BCS_EXCEPTION.
      WRITE: 'Hata Oluştu'(001).
      WRITE: 'Hata türü:'(002), BCS_EXCEPTION->ERROR_TYPE.
      EXIT.

  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*&      Form  REGISTER
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM REGISTER .
  CALL METHOD ALV_0100->REFRESH_TABLE_DISPLAY .
ENDFORM.
