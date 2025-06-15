FROM node:22-bookworm

# 基本的なパッケージのインストール
RUN apt-get update && apt-get install -y \
    git \
    bash \
    curl \
    wget \
    vim \
    nano \
    openssh-client \
    ca-certificates \
    tzdata \
    sudo \
    && rm -rf /var/lib/apt/lists/*

# Google Apps Script開発用のclaspのインストール
RUN npm install -g @google/clasp

# Claude Codeのインストール
RUN npm install -g @anthropic-ai/claude-code

# nodeユーザーの設定
RUN usermod -aG sudo node

# sudoアクセスの設定
RUN echo 'node ALL=(ALL) NOPASSWD:ALL' >> /etc/sudoers

# ワーキングディレクトリの設定
WORKDIR /workspace

# nodeユーザーに切り替え
USER node

# シェルをbashに設定
SHELL ["/bin/bash", "-c"]

# 環境変数の設定
ENV NODE_ENV=development
ENV TZ=Asia/Tokyo 