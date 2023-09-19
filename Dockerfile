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
RUN apt install -y libgdiplus
RUN python3.8 -m pip install --upgrade pip
# End image ubuntu_20.04_py38

# Start image ubuntu_20.04_py39
FROM ubuntu:20.04
ENV DEBIAN_FRONTEND=noninteractive
RUN apt update
RUN echo ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true | debconf-set-selections
RUN apt install -y ttf-mscorefonts-installer
RUN apt install -y python3.9 python3-pip
RUN apt install -y libgdiplus
RUN python3.9 -m pip install --upgrade pip
# End image ubuntu_20.04_py39

# Start image ubuntu_22.04_py310
FROM ubuntu:22.04
RUN apt update && apt install -y python3.10
RUN echo ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true | debconf-set-selections
RUN apt install -y ttf-mscorefonts-installer
RUN apt install -y python3-pip
RUN apt install -y wget
RUN apt install -y libgdiplus
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
RUN apt install -y libgdiplus
RUN python3.11 -m pip install --upgrade pip
RUN wget http://security.ubuntu.com/ubuntu/pool/main/o/openssl/libssl1.1_1.1.0g-2ubuntu4_amd64.deb
RUN dpkg -i ./libssl1.1_1.1.0g-2ubuntu4_amd64.deb
RUN rm -i libssl1.1_1.1.0g-2ubuntu4_amd64.deb
# End image ubuntu_22.04_py311
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
# End image python_3.8_bullseye

# Start image python_3.9_bullseye
FROM python:3.9-bullseye
RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y libgdiplus
RUN python3 -m pip install --upgrade pip
# End image python_3.9_bullseye

# Start image python_3.10_bullseye
FROM python:3.10-bullseye
RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y libgdiplus
RUN python3 -m pip install --upgrade pip
# End image python_3.10_bullseye

# Start image python_3.11_bullseye
FROM python:3.11-bullseye
RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y libgdiplus
RUN python3 -m pip install --upgrade pip
# End image python_3.11_bullseye
#
#--------------------------------------------------------------------------
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
RUN dnf install -y python3-pip libicu libgdiplus
RUN python3 -m pip install --upgrade pip
# End image fedora_32_py38

# Start image fedora_34_py39
FROM fedora:34
RUN dnf install -y curl cabextract xorg-x11-font-utils fontconfig
RUN rpm -i https://downloads.sourceforge.net/project/mscorefonts2/rpms/msttcore-fonts-installer-2.6-1.noarch.rpm
RUN dnf install -y python3-pip libicu libgdiplus
RUN python3 -m pip install --upgrade pip
# End image fedora_34_py39

# Start image fedora_35_py310
FROM fedora:35
RUN dnf install -y curl cabextract xorg-x11-font-utils fontconfig
RUN rpm -i https://downloads.sourceforge.net/project/mscorefonts2/rpms/msttcore-fonts-installer-2.6-1.noarch.rpm
RUN dnf install -y python3-pip libicu libgdiplus
RUN python3 -m pip install --upgrade pip
# End image fedora_35_py310
#
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
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y python3 python3-pip libgdiplus
RUN python3 -m pip install --upgrade pip
# End image debian_py39

# Start image debian_py311
FROM debian:bookworm
RUN apt update
RUN apt install -y wget
RUN wget http://ftp.de.debian.org/debian/pool/contrib/m/msttcorefonts/ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y ./ttf-mscorefonts-installer_3.8_all.deb
RUN apt install -y python3 python3-pip libgdiplus
# End image debian_py311
