# Use an official Node.js runtime as a parent image
FROM node:lts-slim

# Create app directory
WORKDIR /home/node/app

# Install Git and Git Flow
RUN apt-get update && apt-get install -y git git-flow


ENV LD_LIBRARY_PATH=/opt/oracle/instantclient_21_4

RUN apt-get update && \
    apt-get install -y libpq-dev zlib1g-dev build-essential shared-mime-info libaio1 libaio-dev unzip wget --no-install-recommends && \
    wget https://download.oracle.com/otn_software/linux/instantclient/214000/instantclient-sdk-linux.x64-21.4.0.0.0dbru.zip && \
    wget https://download.oracle.com/otn_software/linux/instantclient/214000/instantclient-sqlplus-linux.x64-21.4.0.0.0dbru.zip && \
    wget https://download.oracle.com/otn_software/linux/instantclient/214000/instantclient-basic-linux.x64-21.4.0.0.0dbru.zip && \
    mkdir -p /opt/oracle && \
    cp instantclient-* /opt/oracle/ && \
    cd /opt/oracle/ && \
    unzip instantclient-basic-linux.x64-21.4.0.0.0dbru.zip && \
    unzip instantclient-sdk-linux.x64-21.4.0.0.0dbru.zip && \
    unzip instantclient-sqlplus-linux.x64-21.4.0.0.0dbru.zip && \
    rm -rf /var/lib/apt/lists/* instantclient-basic-linux.x64-21.4.0.0.0dbru.zip instantclient-sdk-linux.x64-21.4.0.0.0dbru.zip instantclient-sqlplus-linux.x64-21.4.0.0.0dbru.zip && \
    apt -y clean && \
    apt -y remove wget unzip && \
    apt -y autoremove && \
    rm -rf /var/cache/apt

# Switch to the "node" user to avoid running as root
USER node

# Define the default command to run (can be overridden)
CMD ["tail", "-f", "/dev/null"]
