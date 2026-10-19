# The WOPI server.
#
# Build: docker build -t wopiserver .
# Run:   docker run -p 8880:8880 -v /path/to/etc-wopi:/etc/wopi:ro -v /path/to/storage:/var/wopi_local_storage wopiserver
#
# /etc/wopi holds wopiserver.conf and the wopisecret and iopsecret files. In the image, the server logs to stdout and
# the local storage is /var/wopi_local_storage.

FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /source
# the SDK pin of global.json is for development: the image builds with the SDK it ships
COPY src/ src/
RUN dotnet publish src/WopiServer/WopiServer.csproj --configuration Release --output /app

FROM mcr.microsoft.com/dotnet/aspnet:10.0
LABEL org.opencontainers.image.title="WopiServer" \
      org.opencontainers.image.description="A WOPI server for ASP.NET Core, ported from the CS3 WOPI server" \
      org.opencontainers.image.source="https://github.com/Papyrine/WopiServer" \
      org.opencontainers.image.licenses="Apache-2.0"
RUN mkdir -p /etc/wopi /var/log/wopi /var/spool/wopirecovery /var/wopi_local_storage && \
    chown $APP_UID /var/log/wopi /var/spool/wopirecovery /var/wopi_local_storage
COPY --from=build /app /app
# container defaults, edited in the shipped defaults so that /etc/wopi can be mounted as a whole
RUN sed -i \
      -e 's|^loghandler = file$|loghandler = stream\nlogdest = stdout|' \
      -e 's|^#storagehomepath = /mnt/your_local_storage$|storagehomepath = /var/wopi_local_storage|' \
      /app/wopiserver.defaults.conf && \
    grep -q '^loghandler = stream$' /app/wopiserver.defaults.conf && \
    grep -q '^storagehomepath = /var/wopi_local_storage$' /app/wopiserver.defaults.conf
USER $APP_UID
EXPOSE 8880
ENTRYPOINT ["dotnet", "/app/wopiserver.dll"]
