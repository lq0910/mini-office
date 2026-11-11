/** 功能：图片选择与剪贴板使用同一真实文件读取及浏览器解码，失败时不修改正文；时间：2026-11-11；作者：lq。 */
import type { NewPicture } from '../upstream/apps/docs/src/renderer/editor/insert-picture';
export async function readWordPictureFile(file: File): Promise<NewPicture> {
  if (!/^image\/(png|jpeg|gif|webp|svg\+xml)$/.test(file.type)) throw new Error('请选择 PNG、JPEG、GIF、WebP 或 SVG 图片。');
  if (!file.size || file.size > 32 * 1024 * 1024) throw new Error('图片必须小于 32 MB 且不能为空。');
  const dataUrl = await new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result)); reader.onerror = () => reject(new Error('图片读取失败')); reader.readAsDataURL(file); });
  const natural = await new Promise<HTMLImageElement>((resolve, reject) => { const img = new Image(); img.onload = () => resolve(img); img.onerror = () => reject(new Error('图片无法解码')); img.src = dataUrl; });
  if (!natural.naturalWidth || !natural.naturalHeight) throw new Error('图片没有有效尺寸。');
  return { dataUrl, base64: dataUrl.slice(dataUrl.indexOf(',') + 1), mime: file.type, widthPx: natural.naturalWidth, heightPx: natural.naturalHeight, label: file.name };
}
