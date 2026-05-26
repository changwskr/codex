# 농협 상호금융 정보계 프로젝트

농협 상호금융 정보계 업무를 위한 Spring Boot 기반 Java 백엔드 프로젝트입니다.

## 기술 스택

- Java 17
- Spring Boot 3.3.6
- Maven
- Spring Web
- Spring Data JPA
- Spring Validation
- Spring Boot Actuator
- H2 Database(local)

## 프로젝트 구조

```text
src
  main
    java/kr/co/nonghyup/sangho/infosystem
      common
        api
        health
    resources
  test
    java/kr/co/nonghyup/sangho/infosystem
```

## 실행

```bash
mvn spring-boot:run
```

기본 프로파일은 `local`이며 서버는 `8080` 포트로 실행됩니다.

## 확인 API

```bash
curl http://localhost:8080/api/v1/system/health
```

## 테스트

```bash
mvn test
```
