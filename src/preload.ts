import { contextBridge } from 'electron';

contextBridge.exposeInMainWorld('geminiDesktop', {
  version: '1.0.0',
});
