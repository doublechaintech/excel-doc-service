all:
	quarkus build --native  --no-tests 
	cp build/excel-doc-service-1.0.0-SNAPSHOT-runner excel-doc-service 
	./excel-doc-service -Xmx64m

dev:
	 ./gradlew --console=plain quarkusDev


docker:
	quarkus build --native  --no-tests -Dquarkus.native.additional-build-args=--initialize-at-run-time=org.apache.poi.util.RandomSingleton
	docker build --no-cache -f src/main/docker/Dockerfile.native -t doublechaintech/excel-doc-service .

run-docker:
	docker run -d --memory 80m --name excel-doc-service  --rm -p 8083:8083  doublechaintech/excel-doc-service 
