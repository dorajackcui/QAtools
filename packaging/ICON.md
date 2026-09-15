# QAtools 应用图标

## 当前资源

- [QAtools-icon.png](QAtools-icon.png)：原始图标母图。暖橙纯色方形背景、米白色 Q，字母采用细颗粒哑光纸质浮雕与柔和短投影。
- [QAtools-icon-rounded.png](QAtools-icon-rounded.png)：正式圆角图标，半径为边长的 18%，四角外侧为真正透明像素；Q、颜色与材质来自原母图。
- [QAtools.ico](QAtools.ico)：Windows 安装程序与可执行文件使用的图标，含 16、20、24、32、40、48、64、128、256 像素。
- [QAtools-icon.svg](QAtools-icon.svg)：保留的平面 Q 轮廓参考；不用于生成当前的材质图标。

图标母图由内置 image_gen 生成，使用用户提供的纸质立体插画作为材质参考、原图标作为 Q 轮廓参考。圆角在资源导出时应用透明蒙版，保持母图内部像素不变；没有采用生成器输出的棋盘格模拟透明图。构建入口仍为 [build_windows_release.ps1](../scripts/build_windows_release.ps1)；更新源资源后需重新构建安装包，已经安装的旧版本不会自动换图标。

从 0.1.6 起，安装程序另外安装带版本号的 ICO，桌面、开始菜单快捷方式和卸载列表直接引用它，避免升级后继续沿用原 EXE 路径的图标缓存。EXE 和安装程序自身仍内嵌同一套图标。

## 导出 ICO

在安装 Pillow 的环境中，从仓库根目录运行：

```powershell
python scripts/export_app_icon.py
```

## 生成提示词

方式：内置 image_gen，未使用 CLI 或独立 API。

第一步使用材质参考图和原 Q 图标：

```text
Use case: style-transfer.
Asset type: final Windows app icon for QAtools, square 1024 x 1024.
Input image 1 is a MATERIAL / LIGHTING STYLE REFERENCE ONLY: the paper-like ivory braces, subtle tactile grain, matte clay / compressed paperboard quality, softly bevelled edges, warm diffused light and gentle shadows. Do not reuse the papers, braces, bars, circles, steps, green background or the composition.
Input image 2 is the EDIT TARGET and identity/layout reference: the existing QAtools Q app icon.
Redesign that icon with the material quality of image 1. Keep exactly one large capital Q, its bold geometric circular ring and short rounded lower-right diagonal tail, comfortably centered. Preserve a readable open counter and a familiar clean Q silhouette at small sizes.
Keep the existing warm orange solid-color rounded-square background (#CC7D5E), with transparent space outside the rounded square. Keep the icon front-facing and upright, no camera tilt or isometric rotation. The background is a flat uniform fill, not a scenery or a gradient; no texture on the background.
Turn the Q into warm ivory matte sculpted paper / dense fine-grained clay, like the cream braces in the reference, with subtle fine paper fibres / tactile grain, shallow three-dimensional depth, gently rounded edges and a soft short lower-right contact shadow. The front face stays light ivory and dominates; modest side depth, not a tall block. Soft upper-left studio light; no gloss or metallic highlights. The Q and its counter must stay crisp and immediately recognizable at 16, 24, 32 and 48 pixels.
Composition: a single finished icon filling the square, rounded-square tile with the same generous corner radius and outer margin as input 2; Q fills roughly 60 percent of the canvas, matching input 2 visual balance.
Text (verbatim): "Q" only.
No extra text, badge, border, ornaments, symbols, platform frame, mockup sheet, palette, sample sizes, watermarks or decorative scene. Preserve true alpha outside the rounded square.
```

第二步保留 Q 材质，修正为完整纯色背景：

```text
Edit the provided QAtools app icon. Preserve the single warm ivory sculpted Q exactly: its silhouette, front-facing geometry, fine matte paper/clay grain, bevels, shallow extrusion, lighting, size and position.
Fix ONLY the background and canvas:
The entire square canvas must be a full-bleed, opaque, SOLID flat warm orange #CC7D5E from edge to edge, all the way into every corner. No rounded-square tile or outer margin: the orange background itself is the entire square icon canvas. Remove every checkerboard square and all surrounding gray pixels. No transparency, no checkerboard pattern, no border, no mockup. Outside the Q's short natural contact shadow, the background is a perfectly flat, untextured single orange color; no vignette or lighting gradient on the background.
This is one finished square app icon artwork. Only the letter Q and the plain orange background. Keep the Q fully visible and centered and retain the tactile ivory material.
```
