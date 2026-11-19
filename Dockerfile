# syntax=docker/dockerfile:1

FROM golang:1.22-alpine AS build
WORKDIR /src

COPY go.mod ./
COPY cmd ./cmd
COPY internal ./internal

RUN CGO_ENABLED=0 GOOS=linux go build -trimpath -ldflags="-s -w" \
    -o /out/xlsx2grand ./cmd/xlsx2grand

FROM alpine:3.20
RUN adduser -D -u 10001 app
COPY --from=build /out/xlsx2grand /usr/local/bin/xlsx2grand
USER app
WORKDIR /data
ENTRYPOINT ["xlsx2grand"]
CMD ["-h"]
