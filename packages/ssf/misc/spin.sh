#!/bin/bash
# spin.sh -- show a spinner (for coverage test)
# Copyright (C) 2014-present  SheetJS

wpid=$1
delay=1
str="|/-\\"
while [ $(ps -a|awk '$1=='$wpid' {print $1}') ]; do
  t=${str#?}
  printf " [%c]" "$str"
  str=$t${str%"$t"}
  sleep $delay
  printf "\b\b\b\b"
done
