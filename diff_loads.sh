#!/bin/sh

./dumpLoad.py SQLite | sort > sqlite.txt && \
	./dumpLoad.py CSV | sort > csv.txt && \
	"$EDITOR" -d csv.txt sqlite.txt

[ "$1" != "-k" ] && \
	[ "$1" != "--keep" ] && \
	rm sqlite.txt csv.txt
