FROM eclipse-temurin:17-jdk
WORKDIR /app
COPY acta-main/ .
RUN javac -cp "libs/*" Servidor.java
EXPOSE 8080
CMD ["java", "-cp", ".:libs/*", "Servidor"]
