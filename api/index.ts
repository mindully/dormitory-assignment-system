import { createApp } from '../server';

let app: any;

export default async function handler(req: any, res: any) {
  if (!app) app = await createApp();
  return app(req, res);
}
