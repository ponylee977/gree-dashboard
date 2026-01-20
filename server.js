/**
 * Gree Dashboard Server
 * 主服务入口 - 包含PPT生成API
 */

import { startServer } from './skills/ppt-generator/api.js';

// 启动服务器
const port = process.env.PORT || 3000;
startServer(port);
