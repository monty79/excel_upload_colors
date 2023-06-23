*&---------------------------------------------------------------------*
*& Report ZUPLOAD_COLORS
*&---------------------------------------------------------------------*
*& Загрузка свойств ячеек
*&---------------------------------------------------------------------*
REPORT zupload_colors.

PARAMETERS p_file TYPE text255.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file .
  DATA retfiletable TYPE filetable.
  DATA retrc TYPE i.
  cl_gui_frontend_services=>file_open_dialog(
      EXPORTING
        multiselection = abap_false
        file_filter = `Excel files (*.XLSX)|*.XLSX`
        default_extension = 'XLSX'
      CHANGING
        file_table = retfiletable
        rc = retrc ).
  READ TABLE retfiletable INTO DATA(lv_file) INDEX 1.
  IF sy-subrc = 0.
    p_file = lv_file.
  ENDIF.

START-OF-SELECTION.

  TRY.
      DATA(lr_excel) = NEW zcl_excel_reader_2007( )->zif_excel_reader~load_file( p_file ).
      DATA(lr_worksheet) = lr_excel->get_active_worksheet( ).
      DATA(highest_column) = lr_worksheet->get_highest_column( ).
      DATA(highest_row)    = lr_worksheet->get_highest_row( ).
      DATA row TYPE i VALUE 1.
      DATA column TYPE i VALUE 1.

      WHILE row <= highest_row.
        WHILE column <= highest_column.
          DATA(col_str) = zcl_excel_common=>convert_column2alpha( column ).
          lr_worksheet->get_cell(
            EXPORTING
              ip_column = col_str
              ip_row    = row
            IMPORTING
              ep_value = DATA(value)
              ep_style = DATA(lr_style)
                ).

          WRITE: / value.
          if lr_style is bound.
            WRITE: 'Фон', lr_style->fill->fgcolor-theme.
          ENDIF.
          free lr_style .
          column = column + 1.
        ENDWHILE.
        WRITE: /.
        column = 1.
        row = row + 1.
      ENDWHILE.

    CATCH zcx_excel INTO DATA(ex).
      DATA(msg) = ex->get_text( ).
      WRITE: / msg.
  ENDTRY.
