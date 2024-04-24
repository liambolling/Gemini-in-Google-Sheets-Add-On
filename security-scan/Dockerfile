# syntax=docker/dockerfile:1
FROM nixos/nix:latest
WORKDIR /usr/scan
COPY . /usr/scan/
RUN mkdir results
RUN nix-env -if https://github.com/fluidattacks/makes/archive/22.09.tar.gz