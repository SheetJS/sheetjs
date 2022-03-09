#!/bin/bash
# make_help.sh -- process listing of targets and special items in Makefile
# Copyright (C) 2016-present  SheetJS
#
# usage in makefile: pipe the output of the following command:
#     @grep -hE '(^[a-zA-Z_-][ a-zA-Z_-]*:.*?|^#[#*])' $(MAKEFILE_LIST)
#
# lines starting with "## " are treated as subtitles
# lines starting with "#* " are treated as plaintext comments
# multiple targets with "## " after the ":" are rendered as separate targets
# if the presumed default target is labeled, it will be assigned a unique color

awk '
BEGIN{recipes=0;}
	!/#[#*] .*$/ {next;}
	{multi=0; isrecipe=0;}
	/^[^#]*:/ {isrecipe=1; ++recipes;}
	/^[^ :]* .*:/ {multi=1}
	multi==0 && isrecipe>0 { if(recipes > 1) print; else print $0, "[default]"; next}
	isrecipe == 0 {print; next}
	multi>0 {
		k=split($0, msg, "##"); m=split($0, a, ":"); n=split(a[1], b, " ");
		for(i=1; i<=n; ++i) print b[i] ":", "##" msg[2], (recipes==1 && i==1 ? "[default]" : "")
	}
END {}
' | if [[ -t 1 ]]; then
awk '
BEGIN {FS = ":.*?## "}
	{color=36}
	/\[default\]/ {color=35}
	NF==1 && /^##/ {color=34}
	NF==1 && /^#\*/ {color=20; $1 = substr($1, 4)}
	{printf "\033[" color "m%-20s\033[0m %s\n", $1, $2;}
END{}' -
else
awk '
BEGIN {FS = ":.*?## "}
	/^#\* / {$1 = substr($1, 4)}
	{printf "%-20s %s\n", $1, $2;}
END{}' -
fi

