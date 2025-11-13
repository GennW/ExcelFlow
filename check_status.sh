#!/bin/bash
while true; do
    if [ -f "output_test.xlsx" ]; then
        echo "Файл output_test.xlsx создан!"
        ls -la output_test.xlsx
        break
    else
        echo "$(date): Файл output_test.xlsx еще не создан"
        sleep 5
    fi
done
