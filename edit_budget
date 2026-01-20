#!/bin/sh

dir=./budgets/
# +/- the number of months to shift from current
delta_months="$1"

if [ -n "$delta_months" ]; then
	[ "$delta_months" -ge 0 ] && delta_months="+$delta_months"
	flag=-"v${delta_months}m"
fi

datestring="$(date $flag +%Y%m)"

$EDITOR "$dir${datestring}budget.csv"
