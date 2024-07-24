#
#----------------------------- UBUNTU--------------------------------
#
# Start image ubuntu_18.04_py36
FROM ubuntu:18.04
RUN apt update
RUN echo ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true | debconf-set-selections
RUN apt install -y ttf-mscorefonts-installer
RUN apt install -y python3.6 python3-pip libgdiplus
RUN python3.6 -m pip install --upgrade pip
# End image ubuntu_18.04_py36

# Start image ubuntu_18.04_py37
FROM ubuntu:18.04
RUN apt update
RUN echo ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true | debconf-set-selections
RUN apt install -y ttf-mscorefonts-installer
RUN apt install -y python3.7 python3-pip libgdiplus
RUN python3.7 -m pip install --upgrade pip
# End image ubuntu_18.04_py37

# Start image ubuntu_20.04_py38
FROM ubuntu:20.04
ENV DEBIAN_FRONTEND=noninteractive
RUN apt update
RUN echo ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true | debconf-set-selections
RUN apt install -y ttf-mscorefonts-installer
RUN apt install -y python3.8 python3-pip
RUN python3.8 -m pip install pillow
RUN python3.8 -m pip install --upgrade pip
# End image ubuntu_20.04_py38

# Start image ubuntu_20.04_py39
FROM ubuntu:20.04
ENV DEBIAN_FRONTEND=noninteractive
RUN apt update
RUN echo ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true | debconf-set-selections
RUN apt install -y ttf-mscorefonts-installer
RUN apt install -y python3.9 python3-pip
RUN python3.9 -m pip install pillow
RUN python3.9 -m pip install --upgrade pip
RUN update-alternatives --install /usr/bin/python3 python3 /usr/bin/python3.8 1
RUN update-alternatives --install /usr/bin/python3 python3 /usr/bin/python3.9 2
RUN update-alternatives --auto python3
# End image ubuntu_20.04_py39

# Start image ubuntu_22.04_py310
FROM ubuntu:22.04
RUN apt update && apt install -y python3.10
RUN echo ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true | debconf-set-selections
RUN apt install -y ttf-mscorefonts-installer
RUN apt install -y python3-pip
RUN apt install -y wget
RUN python3.10 -m pip install pillow
RUN python3.10 -m pip install --upgrade pip
RUN wget http://security.ubuntu.com/ubuntu/pool/main/o/openssl/libssl1.1_1.1.0g-2ubuntu4_amd64.deb
RUN dpkg -i ./libssl1.1_1.1.0g-2ubuntu4_amd64.deb
RUN rm -i libssl1.1_1.1.0g-2ubuntu4_amd64.deb
# End image ubuntu_22.04_py310

# Start image ubuntu_22.04_py311
FROM ubuntu:22.04
RUN apt update && apt install -y python3.11
RUN echo ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true | debconf-set-selections
RUN apt install -y ttf-mscorefonts-installer
RUN apt install -y python3-pip
RUN apt install -y wget
RUN python3.11 -m pip install pillow
RUN python3.11 -m pip install --upgrade pip
RUN update-alternatives --install /usr/bin/python3 python3 /usr/bin/python3.10 1
RUN update-alternatives --install /usr/bin/python3 python3 /usr/bin/python3.11 2
RUN update-alternatives --auto python3
RUN wget http://security.ubuntu.com/ubuntu/pool/main/o/openssl/libssl1.1_1.1.0g-2ubuntu4_amd64.deb
RUN dpkg -i ./libssl1.1_1.1.0g-2ubuntu4_amd64.deb
RUN rm -i libssl1.1_1.1.0g-2ubuntu4_amd64.deb
# End image ubuntu_22.04_py311

# Start image ubuntu_22.04_py312
FROM ubuntu:22.04
ENV DEBIAN_FRONTEND=noninteractive
RUN apt update && apt install -y
RUN apt install software-properties-common -y
RUN add-apt-repository ppa:deadsnakes/ppa -y
Run ln -fs /usr/share/zoneinfo/America/New_York /etc/localtime
Run apt-get install -y --no-install-recommends tzdata
RUN echo ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true | debconf-set-selections
RUN apt install -y ttf-mscorefonts-installer
RUN apt install -y wget
RUN wget http://security.ubuntu.com/ubuntu/pool/main/o/openssl/libssl1.1_1.1.0g-2ubuntu4_amd64.deb
RUN dpkg -i ./libssl1.1_1.1.0g-2ubuntu4_amd64.deb
RUN rm -i libssl1.1_1.1.0g-2ubuntu4_amd64.deb
RUN apt install -y python3.12-full
RUN update-alternatives --install /usr/bin/python3 python3 /usr/bin/python3.10 1
RUN update-alternatives --install /usr/bin/python3 python3 /usr/bin/python3.12 2
RUN update-alternatives --auto python3
RUN python3.12 -m ensurepip --upgrade
RUN python3.12 -m pip install --upgrade setuptools
RUN python3.12 -m pip install --upgrade pip
RUN python3.12 -m pip install pillow
# End image ubuntu_22.04_py312
#
#--------------------------------------------------------------------------
#
#-----------------------------PYTHON BULLSEYE--------------------------------
#
# Start image python_3.7_bullseye
FROM python:3.7-bullseye
RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y libgdiplus
RUN python3 -m pip install --upgrade pip
# End image python_3.7_bullseye

# Start image python_3.8_bullseye
FROM python:3.8-bullseye
RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y libgdiplus
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
# End image python_3.8_bullseye

# Start image python_3.9_bullseye
FROM python:3.9-bullseye
RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
# End image python_3.9_bullseye

# Start image python_3.10_bullseye
FROM python:3.10-bullseye
RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
# End image python_3.10_bullseye

# Start image python_3.11_bullseye
FROM python:3.11-bullseye
RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
# End image python_3.11_bullseye

# Start image python_3.12_bullseye
FROM python:3.12-bullseye
RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
# End image python_3.12_bullseye
#
#--------------------------------------------------------------------------
#
#-----------------------------REDHAT--------------------------------
# Start image redhat_8_py3.8
FROM redhat/ubi8
RUN yum install -y python3.8 libicu openssl
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
COPY ./fonts usr/share/fonts
RUN yum install fontconfig -y
RUN fc-cache -fv
# End image redhat_8_py3.8

# Start image redhat_8_py3.9
FROM redhat/ubi8
RUN yum install -y python3.9 libicu
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
COPY ./fonts usr/share/fonts
RUN yum install fontconfig -y
RUN fc-cache -fv
# End image redhat_8_py3.9

# Start image redhat_8_py3.11
FROM redhat/ubi8
RUN yum install -y python3.11 libicu python3.11-pip
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
COPY ./fonts usr/share/fonts
RUN yum install fontconfig -y
RUN fc-cache -fv
# End image redhat_8_py3.11
#
#-----------------------------ORACLE LINUX --------------------------------------

# Start image oraclelinux_py38
FROM oraclelinux:8.7
RUN yum install -y python3.8 libicu openssl
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
COPY ./fonts usr/share/fonts
RUN yum install fontconfig -y
RUN fc-cache -fv
# End image oraclelinux_py38

# Start image oraclelinux_py39
FROM oraclelinux:8.7
RUN yum install -y python3.9 libicu openssl
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
COPY ./fonts usr/share/fonts
RUN yum install fontconfig -y
RUN fc-cache -fv
# End image oraclelinux_py39

# Start image oraclelinux_py311
FROM oraclelinux:8.7
RUN yum install -y python3.11 python3.11-pip libicu openssl
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
COPY ./fonts usr/share/fonts
RUN yum install fontconfig -y
RUN fc-cache -fv
# End image oraclelinux_py311

# Start image oraclelinux_py312
FROM oraclelinux:8.7
RUN yum install -y python3.12 python3.12-pip libicu openssl
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
COPY ./fonts usr/share/fonts
RUN yum install fontconfig -y
RUN fc-cache -fv
# End image oraclelinux_py312
#
#--------------------------------------------------------------------------
#
#-----------------------------GOOGLE CLOUD-SDK ----------------------------
# Start image cloud_sdk_431_py39
FROM google/cloud-sdk:431.0.0-slim
RUN apt update
RUN apt install -y wget openssl
RUN wget http://archive.ubuntu.com/ubuntu/pool/main/i/icu/libicu66_66.1-2ubuntu2_amd64.deb
RUN dpkg -i ./libicu66_66.1-2ubuntu2_amd64.deb
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
# End image cloud_sdk_431_py39

#
#--------------------------------------------------------------------------
#
#-----------------------------CENTOS ----------------------------
# Start image centos_8_py38
FROM centos:centos8
RUN cd /etc/yum.repos.d/
RUN sed -i 's/mirrorlist/#mirrorlist/g' /etc/yum.repos.d/CentOS-*
RUN sed -i 's|#baseurl=http://mirror.centos.org|baseurl=http://vault.centos.org|g' /etc/yum.repos.d/CentOS-*
RUN yum install -y python3.8 libicu
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
COPY ./fonts usr/share/fonts
RUN yum install fontconfig -y
RUN fc-cache -fv
CMD /bin/bash
# End image centos_8_py38

# Start image centos_8_py39
FROM centos:centos8
RUN cd /etc/yum.repos.d/
RUN sed -i 's/mirrorlist/#mirrorlist/g' /etc/yum.repos.d/CentOS-*
RUN sed -i 's|#baseurl=http://mirror.centos.org|baseurl=http://vault.centos.org|g' /etc/yum.repos.d/CentOS-*
RUN yum install -y python3.9 libicu
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
COPY ./fonts usr/share/fonts
RUN yum install fontconfig -y
RUN fc-cache -fv
CMD /bin/bash
# End image centos_8_py39
#--------------------------------------------------------------------------
#
#-----------------------------FEDORA ----------------------------
# Start image fedora_28_py36
FROM fedora:28
RUN dnf install -y curl cabextract xorg-x11-font-utils fontconfig
RUN rpm -i https://downloads.sourceforge.net/project/mscorefonts2/rpms/msttcore-fonts-installer-2.6-1.noarch.rpm
RUN dnf install -y libicu libstdc++.x86_64 libgdiplus
RUN python3 -m pip install --upgrade pip
# End image fedora_28_py36

# Start image fedora_31_py37
FROM fedora:31
RUN dnf install -y curl cabextract xorg-x11-font-utils fontconfig
RUN rpm -i https://downloads.sourceforge.net/project/mscorefonts2/rpms/msttcore-fonts-installer-2.6-1.noarch.rpm
RUN dnf install -y python3-pip libicu libgdiplus
RUN python3 -m pip install --upgrade pip
# End image fedora_31_py37

# Start image fedora_32_py38
FROM fedora:32
RUN dnf install -y curl cabextract xorg-x11-font-utils fontconfig
RUN rpm -i https://downloads.sourceforge.net/project/mscorefonts2/rpms/msttcore-fonts-installer-2.6-1.noarch.rpm
RUN dnf install -y python3-pip libicu
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
# End image fedora_32_py38

# Start image fedora_34_py39
FROM fedora:34
RUN dnf install -y curl cabextract xorg-x11-font-utils fontconfig
RUN rpm -i https://downloads.sourceforge.net/project/mscorefonts2/rpms/msttcore-fonts-installer-2.6-1.noarch.rpm
RUN dnf install -y python3-pip libicu
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
# End image fedora_34_py39

# Start image fedora_35_py310
FROM fedora:35
RUN dnf install -y curl cabextract xorg-x11-font-utils fontconfig
RUN rpm -i https://downloads.sourceforge.net/project/mscorefonts2/rpms/msttcore-fonts-installer-2.6-1.noarch.rpm
RUN dnf install -y python3-pip libicu
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
# End image fedora_35_py310

# Start image fedora_35_py311
FROM fedora:35
ENV DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=0
RUN yum update -y
RUN yum groupinstall 'Development Tools' -y
RUN yum install -y openssl-devel bzip2-devel libffi-devel sqlite-devel
RUN dnf install -y curl cabextract xorg-x11-font-utils fontconfig
RUN rpm -i https://downloads.sourceforge.net/project/mscorefonts2/rpms/msttcore-fonts-installer-2.6-1.noarch.rpm
RUN dnf install -y libicu wget
RUN wget https://www.python.org/ftp/python/3.11.0/Python-3.11.0.tgz
RUN tar -xf Python-3.11.0.tgz
WORKDIR "/Python-3.11.0"
RUN ./configure --enable-optimizations
RUN make -j 8
RUN make altinstall
RUN python3.11 -m pip install --upgrade pip
RUN update-alternatives --install /usr/bin/python3 python3 /usr/bin/python3.10 1
RUN update-alternatives --install /usr/bin/python3 python3 /usr/local/bin/python3.11 2
RUN update-alternatives --auto python3
RUN sed -i 's|#!/usr/bin/python3|#!/usr/bin/python3.10|g' /usr/bin/dnf
RUN python3 -m pip install pillow
# End image fedora_35_py311

# Start image fedora_35_py312
FROM fedora:35
ENV DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=0
RUN yum update -y
RUN yum groupinstall 'Development Tools' -y
RUN yum install -y openssl-devel bzip2-devel libffi-devel sqlite-devel
RUN dnf install -y curl cabextract xorg-x11-font-utils fontconfig
RUN rpm -i https://downloads.sourceforge.net/project/mscorefonts2/rpms/msttcore-fonts-installer-2.6-1.noarch.rpm
RUN dnf install -y libicu wget
RUN wget https://www.python.org/ftp/python/3.12.0/Python-3.12.0.tgz
RUN tar -xf Python-3.12.0.tgz
WORKDIR "/Python-3.12.0"
RUN ./configure --enable-optimizations
RUN make -j 8
RUN make altinstall
RUN python3.12 -m ensurepip --upgrade
RUN python3.12 -m pip install --upgrade setuptools
RUN python3.12 -m pip install --upgrade pip
RUN update-alternatives --install /usr/bin/python3 python3 /usr/bin/python3.10 1
RUN update-alternatives --install /usr/bin/python3 python3 /usr/local/bin/python3.12 2
RUN update-alternatives --auto python3
RUN sed -i 's|#!/usr/bin/python3|#!/usr/bin/python3.10|g' /usr/bin/dnf
RUN python3 -m pip install pillow
# End image fedora_35_py312
#
#--------------------------------------------------------------------------
#
#-----------------------------AMAZON LINUX --------------------------------

# Start image amazonlinux_py38
FROM amazonlinux:1
RUN yum install -y python38
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
COPY ./fonts usr/share/fonts
RUN yum install fontconfig -y
RUN fc-cache -fv
# End image amazonlinux_py38
#-------------------------------------------------------------------------#
#-----------------------------DEBIAN LINUX --------------------------------
# Start image debian_py37
FROM debian:buster
RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y python3 python3-pip libgdiplus
RUN python3 -m pip install --upgrade pip
# End image debian_py37

# Start image debian_py39
FROM debian:bullseye
RUN apt update
RUN apt install -y wget
RUN apt install -y libicu-dev
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y python3 python3-pip
RUN python3 -m pip install --upgrade pip
RUN python3 -m pip install pillow
# End image debian_py39

# Start image debian_py311
FROM debian:bookworm
RUN apt update && apt upgrade -y
RUN apt install -y wget
RUN wget http://archive.ubuntu.com/ubuntu/pool/main/i/icu/libicu70_70.1-2_amd64.deb
RUN dpkg -i libicu70_70.1-2_amd64.deb
RUN wget http://security.ubuntu.com/ubuntu/pool/main/o/openssl/libssl1.1_1.1.0g-2ubuntu4_amd64.deb
RUN dpkg -i ./libssl1.1_1.1.0g-2ubuntu4_amd64.deb
#RUN apt install -y libicu-dev
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y python3 python3-pip
RUN python3 -m pip install pillow --break-system-packages
# End image debian_py311


# Start image debian_py312
FROM debian:bookworm
ENV DEBIAN_FRONTEND=noninteractive
ENV DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=false
RUN apt update && apt upgrade -y
RUN apt install software-properties-common -y
RUN apt install -y python3-launchpadlib
RUN add-apt-repository ppa:deadsnakes/ppa -y
RUN apt install -y wget
RUN apt install -y libicu-dev
RUN wget http://security.ubuntu.com/ubuntu/pool/main/o/openssl/libssl1.1_1.1.0g-2ubuntu4_amd64.deb
RUN dpkg -i ./libssl1.1_1.1.0g-2ubuntu4_amd64.deb
RUN wget http://archive.ubuntu.com/ubuntu/pool/main/i/icu/libicu66_66.1-2ubuntu2_amd64.deb
RUN dpkg -i ./libicu66_66.1-2ubuntu2_amd64.deb
RUN apt install -y build-essential libssl-dev zlib1g-dev libbz2-dev
RUN apt install -y libreadline-dev libsqlite3-dev curl llvm libncurses5-dev libncursesw5-dev
RUN apt install -y xz-utils tk-dev libffi-dev liblzma-dev python3-openssl git
RUN wget https://www.python.org/ftp/python/3.12.0/Python-3.12.0.tgz
RUN tar -xf Python-3.12.0.tgz
WORKDIR "/Python-3.12.0"
RUN ./configure --enable-optimizations
RUN make -j 8
RUN make altinstall
RUN python3.12 -m ensurepip --upgrade
RUN python3.12 -m pip install --upgrade setuptools
RUN python3.12 -m pip install --upgrade pip
RUN python3.12 -m pip install pillow
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN update-alternatives --install /usr/bin/python3 python3 /usr/bin/python3.11 1
RUN update-alternatives --install /usr/bin/python3 python3 /usr/local/bin/python3.12 2
RUN update-alternatives --auto python3
# End image debian_py312
