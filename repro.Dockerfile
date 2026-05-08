FROM --platform=linux/amd64 ubuntu:24.04
ENV DEBIAN_FRONTEND=noninteractive
RUN apt-get update && apt-get install -y apt-utils dpkg-dev python3
WORKDIR /repo
COPY repro-apt-404.sh /repo/repro-apt-404.sh
CMD ["bash", "/repo/repro-apt-404.sh"]
