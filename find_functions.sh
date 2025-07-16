#!/bin/bash

# Functions to find
functions=(
    "getSpreadsheetData"
    "getSpreadsheetDataPaginated" 
    "updateMachineStatus"
    "updateMultipleStatuses"
    "getLocationSheetData"
    "updateLocationSheetCell"
    "getDestinationSheetData"
    "getDestinationSheets"
    "checkDataConsistency"
)

echo "=== 関数の位置と範囲 ==="
echo ""

for func in "${functions[@]}"; do
    # Find function start line
    start_line=$(grep -n "^function $func\s*(" Code.gs | cut -d: -f1)
    
    if [ -n "$start_line" ]; then
        # Find next function or end of file
        next_func_line=$(awk "NR>$start_line && /^function/ {print NR; exit}" Code.gs)
        
        if [ -n "$next_func_line" ]; then
            end_line=$((next_func_line - 1))
        else
            end_line=$(wc -l < Code.gs)
        fi
        
        echo "### $func"
        echo "開始行: $start_line"
        echo "終了行: $end_line"
        echo "範囲: $start_line-$end_line"
        echo ""
    else
        echo "### $func"
        echo "見つかりませんでした"
        echo ""
    fi
done