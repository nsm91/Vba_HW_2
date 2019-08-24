Attribute VB_Name = "Module1"
Sub StockVol():

Dim last_row As Integer
Dim last_row2 As Integer
Dim vol_total As Double
Dim export_table_counter As Integer
Dim date_min As Double
Dim open1 As Double
Dim close1 As Double
Dim per_chg As Double
Dim abs_chg As Double

date_min = 2
export_table_counter = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Cells(1, 9).Value = "Stock Name"
Cells(1, 10).Value = "Total Volume"
Cells(1, 11).Value = "Absolute Change"
Cells(1, 12).Value = "Percent Change"


For Stock = 2 To lastrow
    vol_total = 0
    If Cells(Stock + 1, 1).Value <> Cells(Stock, 1).Value Then
        vol_total = vol_total + Cells(Stock, 7).Value
        Cells(export_table_counter, 10).Value = vol_total
        Cells(export_table_counter, 9).Value = Cells(Stock, 1).Value
        'Attempt to integrate counter to calculate range for dates
        open1 = Cells(date_min, 3).Value
        close1 = Cells(Stock, 6).Value
        per_chg = (open1 - close1) / open1
        abs_chg = open1 - close1
        Cells(export_table_counter, 11).Value = abs_chg
        Cells(export_table_counter, 12).Value = per_chg
        If per_chg > 0 Then
            Cells(export_table_counter, 11).Interior.ColorIndex = 35
        Else
            Cells(export_table_counter, 11).Interior.ColorIndex = 46
        End If
        date_min = Stock + 1
        export_table_counter = export_table_counter + 1
        
    Else
        vol_total = vol_total + Cells(Stock, 7).Value

    End If
Next Stock

'Part 3 interate through summary table for largest value
lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
Dim max_vol As Double
Dim max_abs_chg As Double
Dim max_per_chg As Double
max_vol = 0
max_abs_chg = 0
max_per_chg = 0

For j = 2 To lastrow2
    If Cells(j, 10).Value > max_vol Then
        max_vol = Cells(j, 10).Value
    End If
    If Cells(j, 11).Value > max_abs_chg Then
        max_abs_chg = Cells(j, 11).Value
    End If
    If Cells(j, 12).Value > max_per_chg Then
        max_per_chg = Cells(j, 12).Value
    End If
Next j

Cells(1, 14).Value = "Max Volume"
Cells(1, 15).Value = "Max Absolute Change"
Cells(1, 16).Value = "Max Percent Change"
Cells(2, 14).Value = max_vol
Cells(2, 15).Value = max_abs_chg
Cells(2, 16).Value = max_per_chg

End Sub
