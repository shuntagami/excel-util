const env = process.env

export const APP_ENV = env['APP_ENV']

export const X_API_KEY = env['X_API_KEY'] as string

export const BLUEPRINT_API_BASE_URL = env['BLUEPRINT_API_BASE_URL'] as string

export const AWS_REGION = env['AWS_REGION'] as string

export const S3_BUCKET_NAME = env['S3_BUCKET_NAME'] as string

export const S3_ENDPOINT = env['S3_ENDPOINT']

export const AWS_ACCESS_KEY_ID = env['AWS_ACCESS_KEY_ID']

export const AWS_SECRET_ACCESS_KEY = env['AWS_SECRET_ACCESS_KEY ']
