# syntax = docker/dockerfile:1.4

FROM node:20-slim

WORKDIR /usr/src/app

COPY package*.json ./

RUN npm ci

# lambdaは以下をやる必要がありそう
# why: https://github.com/lovell/sharp/issues/3717
# ref: https://sharp.pixelplumbing.com/install#aws-lambda
RUN rm -rf node_modules/sharp
RUN SHARP_IGNORE_GLOBAL_LIBVIPS=1 npm install --arch=x64 --platform=linux --libc=glibc sharp

COPY . .

RUN npm run build

CMD ["npm", "run", "dev"]
