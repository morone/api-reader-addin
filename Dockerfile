FROM node:latest

ADD . /code
WORKDIR /code

RUN npm install

CMD ["npm", "start", "--host=0.0.0.0"]