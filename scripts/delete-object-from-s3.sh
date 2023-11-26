#!/usr/bin/env bash

BUCKET="shuntagami-demo-data"

awslocal s3 rm "s3://${BUCKET}" --recursive
