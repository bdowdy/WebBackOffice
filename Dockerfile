# Build stage
FROM mcr.microsoft.com/dotnet/sdk:9.0 AS build
WORKDIR /src
COPY MSaleWebServer.csproj .
RUN dotnet restore MSaleWebServer.csproj
COPY . .
RUN dotnet publish MSaleWebServer.csproj -c Release -o /app/publish /p:UseAppHost=false

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:9.0 AS runtime
WORKDIR /app
COPY --from=build /app/publish .
# Railway injects PORT at runtime; default to 8080 for local runs
CMD ASPNETCORE_URLS=http://0.0.0.0:${PORT:-8080} dotnet MSaleWebServer.dll
