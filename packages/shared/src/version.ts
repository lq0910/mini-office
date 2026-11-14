import { version } from '../../../package.json';

/** 功能：从根包读取当前版本，统一首页、设置和关于页的版本展示；时间：2026-10-24；作者：lq。 */
export const APP_VERSION = version;

/** 功能：关于页显示独立构建号，区分同一版本下的本地测试包；时间：2026-11-14；作者：lq。 */
export const APP_BUILD_ID: string = import.meta.env.VITE_BUILD_ID ?? 'dev';
