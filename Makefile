build:
	docker build -t fivespace_gen:0.1 .

run:
	@docker rm -f test
	docker run -it --name test -v ${PWD}:/usr/src/app fivespace_gen:0.1 /bin/bash -c "npm config set registry https://registry.npmmirror.comn;npm i; npm run start"
