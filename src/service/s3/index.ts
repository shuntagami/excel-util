import { S3Client } from '@aws-sdk/client-s3'
import {
  APP_ENV,
  AWS_ACCESS_KEY_ID,
  AWS_REGION,
  AWS_SECRET_ACCESS_KEY,
  S3_ENDPOINT
} from './../../constant/constant'
import type { StorageService } from '../../interface/storageService'
import { S3Service } from './s3Service'
import { MockS3Service } from './mockS3Service'

let storageService: StorageService

if (APP_ENV === 'local') {
  storageService = new MockS3Service()
} else {
  let credentials
  if (AWS_ACCESS_KEY_ID === undefined || AWS_SECRET_ACCESS_KEY === undefined) {
    credentials = {
      accessKeyId: '',
      secretAccessKey: ''
    }
  } else {
    credentials = {
      accessKeyId: AWS_ACCESS_KEY_ID,
      secretAccessKey: AWS_SECRET_ACCESS_KEY
    }
  }

  const client = new S3Client({
    credentials,
    region: AWS_REGION,
    endpoint: S3_ENDPOINT,
    forcePathStyle: true
  })

  storageService = new S3Service(client)
}

export { storageService }
