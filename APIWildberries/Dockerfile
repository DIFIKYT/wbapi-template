FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY . .
RUN dotnet publish -c Release -r linux-x64 --self-contained true -o /app

FROM mcr.microsoft.com/dotnet/runtime-deps:8.0 AS runtime

RUN apt-get update && apt-get install -y locales \
    && locale-gen ru_RU.UTF-8

ENV LANG=ru_RU.UTF-8
ENV LANGUAGE=ru_RU:ru
ENV LC_ALL=ru_RU.UTF-8

WORKDIR /app
COPY --from=build /app .
ENTRYPOINT ["./APIWildberries"]
