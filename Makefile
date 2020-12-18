
DATE=`date +%Y%m%d%H%M%S`

.PHONY: all
all: commit-m

.PHONY: add
add: 
	git add --all .

# Commiting with a timestamp for a message
.PHONY: commit-m
commit-m: add
	git commit -m $(DATE)

# Commiting with custom notification message
.PHONY: commit
commit: add
	git commit

.PHONY: push
push:
	git push -u origin master

.PHONY: pull
pull:
	git pull

.PHONY: status
status:
	git status

.PHONY: openall
openall:
	geany *js &

.PHONY: help
help:
	@echo "make [all|add|commit-m|commit|push|pull|status|openall|help]"
