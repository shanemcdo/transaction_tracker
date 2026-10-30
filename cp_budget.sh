#!/bin/bash

dir=./budgets/
now=`date +%Y%m`
next=`date -v+1m +%Y%m`


cp -v "$dir${now}budget.csv" "$dir${next}budget.csv"
