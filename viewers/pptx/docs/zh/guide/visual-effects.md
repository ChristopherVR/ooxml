---
title: 视觉效果还原
description: 具体 PowerPoint 视觉效果在 CSS/SVG 中的实现方式，以及每项效果对应的 COM 实测依据。
---

# 视觉效果还原 {#visual-effect-fidelity}

幻灯片渲染为 HTML/CSS/SVG，而非 Canvas，以便在任意缩放下保持文字清晰，支持原生无障碍和 DOM 交互。因此，少数没有精确 CSS/SVG 对应实现的 PowerPoint 效果，会采用经过选择的专门技术重现。本页记录已对照真实 PowerPoint 输出验证、效果可靠的实现；尚未解决的差距请参见[已知限制](/zh/guide/limitations)。

## 倒影（`a:effectLst/a:reflection`） {#reflections-a-effectlst-a-reflection}

倒影渲染为镜像的同级节点，反映元素完整的渲染内容：对于形状或图片，包含填充、轮廓和自身文本；对于组合，递归包含每个子元素及其填充、轮廓和文本。五种绑定均支持 `@sx` / `@sy` / `@kx` / `@ky` / `@rot` / `@fadeDir` / `@algn`。

即使组合没有自己的填充，也会渲染组合自身的倒影，以及阴影、发光和柔化边缘。这些效果通过 CSS `filter` 应用到组合的合成栅格结果上，不使用 `box-shadow`，因为后者会给组合的边界矩形而非实际内容添加阴影。如果带倒影的组合内部，某个子元素也有自己的倒影，就会发生双重镜像，与 PowerPoint 的合成方式一致：组合倒影基于已经完整渲染的组合内容，其中包含子元素自身的倒影。

## 柔化边缘（`a:softEdge`） {#soft-edges-a-softedge}

使用 SVG 滤镜，仅羽化形状的 alpha 边缘：先腐蚀，再模糊，最后与原始填充合成。内部填充和文字保持清晰，不会模糊整个元素。

**根据 COM 实测修正（2026-09-10）。** 原实现将文档中的 `@rad` 直接作为单次 `feGaussianBlur(stdDeviation = rad)` 的参数，再通过 `in` 与源图形合成。这样得到的过渡形状和宽度都不正确：硬边经过高斯模糊后，在原始边界处恰好为 50% 不透明度，向内约 3 倍半径才达到完全不透明。对真实 PowerPoint 2016 进行测量：使用 `Slide.Export` 导出 PNG，幻灯片尺寸为 1280×720，一个矩形设置 `a:softEdge rad="190500"`，即 20px，第二个样例设置 `rad="381000"`，即 40px；沿穿过形状边缘的水平扫描线采样，结果如下：

| 设置的半径 | 边界处不透明度 | 约 50% 不透明度               | 约 100%（饱和）              |
| ---------- | -------------- | ----------------------------- | ---------------------------- |
| 20px       | 约 0%          | 向内约 18px（半径的 0.9 倍）  | 向内约 34px（半径的 1.7 倍） |
| 40px       | 约 0%          | 向内约 35px（半径的 0.88 倍） | 向内约 72px（半径的 1.8 倍） |

两种半径都证实，PowerPoint 的羽化效果在原始边界处接近透明，而非 50%，在向内约 1.75 至 1.8 倍设置半径处达到完全不透明，而非原始高斯模糊所产生的约 3 倍；该关系随设置半径线性变化。现在渲染器先将 alpha 向内腐蚀 `0.9 x radius`，再使用更窄的模糊（`0.3 x radius`，其自身的三倍标准差扩散范围约为半径的 0.85 倍）羽化，与两条实测曲线的偏差都约为一个像素。实现统一位于 `src/ui/src/pptx/render/visual-effects.ts` 的 `getSoftEdgeSvgFilter`，五种绑定以相同方式使用。

## 路径渐变（`a:gradFill/a:path`） {#path-gradients-a-gradfill-a-path}

`circle` 和 `shape` 路径类型使用原生的椭圆或圆形 CSS/SVG 径向渐变渲染。PowerPoint 对这两种类型的渲染本身就是椭圆形，因此无需近似处理。

`rect` 路径类型不同：PowerPoint 沿形状自身的边界矩形渐变，等值线是带直角的同心矩形，即切比雪夫距离或 L∞ 距离场。原生 CSS 和 SVG 径向渐变都是椭圆形，无法生成这种效果。`src/ui/src/pptx/render/path-gradient-rect.ts` 通过堆叠 40 层与坐标轴对齐的嵌套 `<rect>` 色带，直接渲染真实矩形距离场，先绘制最大层，最后绘制最小层。

**COM 实测（2026-09-10）。** 在 PowerPoint 2016 中，用 `Slide.Export` 将包含 `a:gradFill path="rect"` 的样例导出为 PNG：渐变色标为红、绿、蓝，`a:fillToRect` 向内缩进 30%。在整个渐变宽度上每隔 5px 采样，颜色变化完全平滑连续，没有可见色带；内侧纯色区域也准确落在 `fillToRect` 指定的位置。PowerPoint 自身的渲染没有需要模仿的离散分层；这里的 40 层离散色带是在不存在原生直角渐变图元时作出的取舍。即使渐变覆盖整个宽度，40 步中每条可见色带也只有几个 CSS 像素，在正常缩放下不可察觉，实测样例的屏幕输出也确认了这一点。底层距离场形状，包括沿两轴的线性变化、对称衰减和矩形等值线，与 PowerPoint 实测颜色值完全一致。

## 像素化切换滤镜（`p:animEffect/@filter="pixelate"`） {#pixelate-transition-filter-p-animeffect-filter-pixelate}

`p:animEffect/@filter` 可以指定的每个 SMIL 滤镜族，都映射为与 PowerPoint 播放一致的显现或隐藏效果。`pixelate` 是其中唯一直接跳到结束状态、而非渐进过渡的滤镜族，因为 PowerPoint 自身就是这样显示的。

**COM 实测（`e2e/pptx/fixtures/pixelate-filter.pptx`，PowerPoint 2016，`Presentation.CreateVideo`，并与仅将滤镜替换为 `filter="dissolve"`、其余字节相同的对照文稿逐帧比较）。** `dissolve` 对照文稿会明显溶解：显现中途的早期帧与稍后稳定帧存在逐像素差异。`pixelate` 文稿则没有变化：其点击步骤最早和最晚的帧逐字节相同，目标形状从该步骤的第一帧开始就已经完全绘制，且完全不透明。PowerPoint 对 `filter="pixelate"` 不执行任何动画，处理方式与无法解释的构建效果相同：静默跳到解析后的结束状态。`pixelate` 是模式允许的 `ST_TransitionFilterType` 值（ECMA-376 20.1.8.49），但真实 PowerPoint 没有对应的宿主实现。

因此，渲染器默认将 `pixelate` 解析为与 `p:animEffect/@filter="cut"` 相同的 `cutIn` / `cutOut` 关键帧，即真正的瞬间切换。在线运行也已验证：`e2e/pptx/animation-pixelate-filter.spec.ts` 断言进入效果播放 `pptx-cutIn`，而非渐进显现。这样与 PowerPoint 行为保持一致，不额外播放 PowerPoint 从未显示的动画。仍然提供显式、默认关闭的块状马赛克显现选项 `pixelateMosaicAnimation`，位于“文件 > 选项 > 高级 > 幻灯片放映 > 为像素化切换显示马赛克效果”，供希望显示动画的查看器使用。其实现 `src/ui/src/pptx/render/animation-pixelate-filter.ts` 使用自包含 SVG `<filter>` data URI，通过离散 `@keyframes` 停点切换，每个可见块显示元素自身的真实内容。决策统一实现在 `src/ui/src/pptx/render/animation-filter-effects.ts` 的 `resolveFilterEffect` 中，经 `PresentationAnimationController.fromSlide` 的 `pixelateMosaic` 选项传递，五种绑定以相同方式使用。

## 电影式切换（`cube`、`box`、`flip`、`rotate`、`pageCurl`、`origami` 等） {#cinematic-transitions-cube-box-flip-rotate-pagecurl-origami}

这一组效果在二维幻灯片图层上通过 CSS 关键帧实现透视、旋转和卷曲，而非真正的体积三维渲染。这是有意的设计：每种预设都有独立调校的关键帧，重现 PowerPoint 在屏幕上的运动形态，而不是统一使用通用三维场景。

已对照 PowerPoint 2016 的 `CreateVideo` 帧进行 **COM 实测**：

- `cube` / `rotate` 共用贴齐屏幕的铰链旋转方式；`cube` 使用基础 `<p14:prism>`，`rotate` 在同一元素上设置 `isContent="1"`。
- `box` / `orbit` 共用向深处退缩的铰链效果，产生间隙，并在两个轴上形成透视缩短；`box` 设置 `isInverted="1"`，`orbit` 设置 `isContent="1" isInverted="1"`。
- `doors` / `window` 的 `horz` 沿上下方向打开。
- `fallOver` 让离开的幻灯片绕顶部铰链倾倒，而非让进入的幻灯片倾倒。
- `reveal` 在前半段保持真实的暗色间隙，然后进入的幻灯片才淡入。
- `warp` 是径向缩放模糊爆发，而不是倾斜。
- `crush` 向中心揉皱，而不是平面的垂直压缩。
- `flythrough`、`gallery`、`ferris`、`conveyor`、`switch`、`pageCurl`（单页或双页）、`peelOff`、`drape`、`ripple`、`flash`、`zoom` 和 `origami` 同样经过 COM 实测，运动形态与关键帧一致。
- COM 已确认，`vortex`、`honeycomb`、`glitter`、`shred`、`fracture`、`curtains` 和 `airplane` 在真实 PowerPoint 中表现为多个独立碎片、图块或粒子，`airplane` 则是真正的纸飞机轮廓折叠。现在每种效果都基于同一测量结果，渲染数量受限的独立片段，每个片段有自己的 clip-path、变换和透明度动画，而不是单个平面图层。实现位于 `src/ui/src/pptx/render/slide-transition-fragments.ts` 的 `getFragmentedTransitionDescriptor`。

## 三维形状与场景（`a:sp3d` / `a:scene3d`） {#_3-d-shapes-and-scenes-a-sp3d-a-scene3d}

相机预设映射为真实的 CSS 变换（`transform-style: preserve-3d`）。对于 `perspective*` 和 `isometric*` 预设，使用 COM 实测的精确 `matrix3d` 单应变换：对每种预设的投影四边形做凸包拟合，归一化到单位正方形，再按元素渲染尺寸缩放，而不是近似旋转。单轴 `perspectiveAbove` / `Below` / `Left` / `Right` 族是纯缩放加偏移，没有梯形畸变；双轴 `*Facing` / `Contrasting*` / `Heroic*` 族的离轴消失点也得到精确表示。COM 已确认，`oblique*` / `legacyOblique*` / `legacyPerspective*` / `orthographicFront` 会使形状正面保持平面；只有拉伸形状的侧面响应这些旧式拉伸方向相机。

显式的 `a:camera/a:rot` / `@fov` / `@zoom` 覆盖通过通用参数化相机函数（`visual-3d-camera-parametric.ts`），构建与预设表同类的精确单应变换。拉伸深度通过真实的 `translateZ` 侧面渲染；每种预设显示哪一侧（上、下、左、右）都经过 COM 实测，共 44 种预设，使用边缘带像素分析。轮廓按文档指定宽度渲染为真实的实心 CSS outline 环。

棱台预设通过真实的 SVG 光照 `<filter>`（`visual-3d-bevel-lighting.ts`）渲染：根据形状自身的 alpha 轮廓构建高度图，在 `feDistantLight` 下使用 `feDiffuseLighting` 和 `feSpecularLighting` 照明，不是 CSS `box-shadow` 近似。材质预设（`a:sp3d/@prstMaterial`）参与同一光照计算，而非独立 CSS 滤镜；`a:lightRig/@rig` 的仰角表已针对全部 27 个 `ST_LightRigType` 值进行 COM 校准。场景地平面（`a:backdrop`）有意不渲染自身阴影，与 PowerPoint 一致：未设置阴影效果时没有地平面视觉痕迹；在倾斜背景上存在阴影时，其形状也无法用 CSS 表示。艺术字和文本体的三维相机预设（`a:bodyPr/a:scene3d`）调用与形状级相同的相机函数。

**已解决（2026-09-16）。** 下文提到的 `relaxedInset` / `slope` / `hardEdge` 棱台剖面差距已修复：新增的非单调高度重映射阶段，可以复现实测的“亮峰后接暗谷”横截面。详见下面的专门小节。

**已解决（2026-09-24）。** 重新校准灯光组仰角后，暴露了镜面与漫反射耦合的缺陷：`metal` 的镜面反射常量按过去假定的 `threePt` 45 度仰角调校，而实际测得 `threePt` 为 74 度，`flat` / `contrasting` 族为 90 度，在这些仰角下形状平坦内部会饱和为纯白。尝试过两种修正，但都因回归而放弃：将镜面光源固定在较低的独立仰角，使 `angle` / `softRound` 变差；将镜面贡献遮罩到弯曲棱台带，则使所有剖面变差。最终采用的修复是：`MaterialLighting.specularElevationCapDeg`（`visual-3d-bevel-lighting-material.ts`）让镜面 `<feDistantLight>` 拥有自己的仰角 `Math.min(rigElevationDeg, cap)`，与漫反射光源的仰角相互独立，并针对 COM 基准值，联合重新拟合 `metal` 的 `specularConstant` / `specularExponent`（`surfaceScaleMultiplier` 保持不变，因为降低它会同时压平漫反射带，那是另一个问题）。使用无界面 Chromium 渲染真实滤镜链，并与 134 种真实 PowerPoint `Slide.Export` 条件对比（全部 27 个 `a:lightRig/@rig` 值 × `circle` / `angle`，加上原有的 8 个灯光组 × 12 种剖面，`dir="t"`，完整的 96 点横截面平均绝对误差，而不只是 2 个边缘点）：

|                          | 平均绝对误差（0-255） |
| ------------------------ | --------------------- |
| 仰角耦合（修复前）       | 75.0                  |
| 解耦并重新拟合（修复后） | 36.4                  |

在上限之下，该截断不起作用（`morning`、`sunrise` 以及其余本就仰角较低的灯光组不受影响），并且只有 `metal` 带有上限，因此其他材质的滤镜标记没有变化。

### 非单调棱台剖面高度重映射（`relaxedInset` / `slope` / `hardEdge`） {#non-monotonic-bevel-profile-height-remap-relaxedinset-slope-hardedge}

**2026-09-16 已解决。** COM 横截面呈现真实“亮峰后接暗谷”双重过渡的三种剖面（见上文的剖面形状测量），现在使用专门的重塑阶段，而不再强迫共享的模糊（加腐蚀）高度坡道去近似它。`visual-3d-bevel-lighting-profile.ts` 中 `BEVEL_PROFILE_HEIGHT_MAP` 针对这三种剖面的条目带有 `heightTransferTable`：一个有序的停点数组，送入 `feComponentTransfer` / `feFuncA type="table"` 图元，插在现有的 `feGaussianBlur`（加 `feMorphology` 腐蚀）高度坡道与 `feDiffuseLighting` / `feSpecularLighting` 之间（`visual-3d-bevel-lighting.ts` 的 `renderLayerPrimitives`）。由于 `feFuncA` 逐点重映射高度值，非单调的表会将高度场自身的空间导数（也就是被照亮的表面法线）重塑为真正的双峰剖面，即使底层被重映射的模糊 / 腐蚀坡道仍然是单调的。这不需要新的几何图元，只需要一个重塑阶段。

每种剖面的表就是该剖面自己固定下来的 24pt COM 横截面曲线（`visual-3d-bevel-lighting-tables.test.ts` 的 `MEASURED_24PT_CURVES`，从上边缘向内取 10 个点），并通过 `(brightness - min) / (max - min)` 归一化到 `[0, 1]`：

| 剖面           | 归一化表（10 个停点，边缘 -> 内部）                                      |
| -------------- | ------------------------------------------------------------------------ |
| `relaxedInset` | 0.671, 0.958, **1.0**, 0.975, **0.0**, 0.320, 0.839, 0.924, 0.924, 0.924 |
| `slope`        | 0.519, **1.0**, 0.967, 0.960, **0.0**, 0.426, 0.100, 0.960, 0.960, 0.960 |
| `hardEdge`     | 0.279, **1.0**, 0.938, 0.938, **0.0**, 0.040, 0.793, 0.938, 0.938, 0.938 |

**这是什么，不是什么。** 这是有意采用并已记录的近似，并不声称是反演的光照模型：对于曲率平缓、以漫反射为主的高度场，渲染亮度几乎直接跟随局部高度，因此将实测亮度曲线自身的形状用作高度重映射曲线，就能较准确地复现实测的峰、谷和恢复位置，而无需反向求解漫反射和镜面反射方程。验证分为两个方面：

- **结构上**（`visual-3d-bevel-lighting-tables.test.ts`、`visual-3d-bevel-lighting.test.ts`）：每个表都是真正的非单调序列（在最大值处有内部峰，随后是明显的谷，在该表自身 0 到 1 的刻度上至少低 0.3，然后恢复）；每个表都直接由该文件固定的 `MEASURED_24PT_CURVES` 推导而来（回归测试会重新计算归一化，若固定数据与表发生偏离就会失败）；生成的 SVG 滤镜标记只对这三种剖面带有送入光照图元的新 `feComponentTransfer` / `feFuncA` 阶段，其余 9 种剖面（即使没有这一阶段，与实测曲线的 RMSE 也已在 1 至 7 以内）保持不变。
- **没有**针对新的 COM 渲染重新验证：此次变更没有进行新的 `Slide.Export` 测量。这里重复使用的是 2026-09 已有的固定数据作为基准，因此精确的屏幕亮度应视为针对已测数据的分析 / 几何拟合，而不是像上文剖面形状测量那样引用的、经过新像素验证的 RMSE 数值。

### 测量来源 {#measurement-provenance}

- 相机单应变换：后续的 27 点 `lat` × `lon` × `rev` COM 网格测量（2026-09），将投影中的组合轴项替换为精确的旋转组合交叉项。27 个点的平均角点误差降至 0.61%，最大 2.1%，消除了先前每种情况只测量一次时留下的约 25% 至 29% 残差。`lon` 的符号已由 COM 直接确认；单轴情况无法观察 `lat` 的符号，因为 `cos` 是偶函数，但 27 点网格已联合 `lon` 确认它。`rev` 的符号通过单独的 45 度滚转 COM 样例隔离确认，4 个角点的方向偏差约在 10 度以内。
- 拉伸侧面：针对单应变换覆盖的 44 种预设，使用边缘带像素分析测量。其中 37 种具有可见侧面的预设，其侧面投影形状已通过 COM 验证：后边缘是前边缘沿实测的、随深度线性变化的屏幕空间向量平移得到。PowerPoint 中，`perspectiveFront`、`orthographicFront`、`perspectiveLeft` / `Right`、`obliqueLeft` / `Right` 和 `perspectiveHeroicExtremeLeftFacing` 完全不显示侧面。
- 棱台高光方向：12 种 `a:bevelT` 剖面中，有 9 种经 COM 确认为将 `a:lightRig/@dir` 吸附到上下左右主方向；`softRound` 测得的是相反主方向的边缘；`slope` / `hardEdge` 没有清晰的方向性亮度信号（参见上面的差距 1）。
- 棱台高度图形状：2026-09 的测量将真实横截面亮度曲线（40 个点、6pt 和 24pt 棱台、全部 12 种剖面）与滤镜已有的 `blurFactor` / `morphologyFactor` / `surfaceScaleFactor` 参数拟合。12 种中有 9 种得到较接近的结果：`circle`、`convex`、`softRound`、`divot`、`angle`、`cross`、`coolSlant`、`riblet`、`artDeco`。完整表格和固定曲线见 `visual-3d-bevel-lighting-tables.ts` 的 `BEVEL_PROFILE_HEIGHT_MAP` 文档注释及其 `.test.ts`。其余 3 种（`relaxedInset` / `slope` / `hardEdge`）已于 2026-09-16 通过 `heightTransferTable` 重塑阶段解决，该阶段直接由这些相同的固定曲线拟合（没有新的 COM 渲染）；详见上面的专门小节。
- 材质：对 `circle` / `angle` / `hardEdge` / `softRound`、4 个主方向、`matte` / `metal` 共 32 种条件进行 COM 测量，在距每条边 0.15 英寸处采样。使用 SVG 滤镜后，`matte` 平均绝对亮度误差从 box-shadow 基线的 56.3 降至 34.8，`metal` 的原始基线为 61.2。`metal` 需要重新校准，提高 `diffuseConstant`、降低 `surfaceScale` 倍率，才能使 `angle` / `hardEdge` / `softRound` 在每个方向上达到或优于基线。只有 `metal` / `circle` 仍难以拟合，因为 COM 测得的阴影侧反射率超过 `feDiffuseLighting` 的 `N.L<=0` 截断机制所能产生的范围。重新拟合 `circle` 自身高度图后，误差才降至 39.5，优于 54.9 基线，并通过新的 COM 测量及实际滤镜的无界面 Chromium 渲染确认。同样重新检查了 `angle` / `hardEdge` / `softRound`，误差为 41.5 / 62.0 / 27.5，对应基线约为 70 / 69.5 / 50。
- 灯光组仰角：测量了全部 27 个 `ST_LightRigType` 值，使用中灰色 `matte` / `circle` 正方形，在距边缘 0.15 英寸处采样；独立的 `dir="t"` / `dir="r"` 测量差异小于 1 个亮度单位。十个灯光组（`harsh`、`balanced`、`twoPt`、各 `legacy{Flat,Normal,Harsh}` 族的 `*4` 变体，以及 5 个天气灯光组）在 `a:lightRig/@dir` 的相反边缘出现高光。4 个带编号的旧式变体不能互换：在该采样偏移处，`*1` / `*3` 没有方向性，`*2` / `*4` 是方向镜像。另一轮测量对中灰色 `circle` 棱台正方形的正中心采样，覆盖 `matte` / `metal` / `plastic` 和 15 个灯光组值，仰角为 54 至 90 度。三种材质在每个灯光组下都逐像素相同，确认在所有仰角下，平坦法线处的材质响应均为零，而不只是在低仰角下如此。
- 轮廓，以及高光和阴影带宽度随棱台深度（6pt 对比 24pt）的缩放关系，均已通过 COM 确认；`extrusionClr` / `contourClr` 颜色得到遵循。背景阴影使用同时带有真实 `a:outerShdw` 和 scene3d backdrop 的形状测量：PowerPoint 在没有阴影效果时不显示地平面痕迹；即使有阴影，水平或近水平背景也与无背景无法区分，而大幅倾斜背景会投影出非凸的剪切形状，无法用 CSS `box-shadow` 表示。
- 艺术字和文本体相机：`perspectiveHeroicLeftFacing` 和 `perspectiveLeft` 已通过 COM 验证，使用填充字形“墨迹”矩形的凸包，结果与形状测量偏差约在 1% 以内；`isometricTopUp` 已视觉确认，但其 45 度滚转使自动角点拟合不够可靠，因此未给出百分比。

## 艺术字包络字形轮廓变形（`a:prstTxWarp` inflate/deflate/can） {#wordart-envelope-glyph-outline-warping-a-prsttxwarp-inflate-deflate-can}

艺术字预设中的 `inflate` / `deflate` / `can`“包络”族，在独立的上下两条曲线之间弯曲文本，因此字形高度（而不只是基线）会随水平位置变化。两种技术配合完成渲染，并根据每个字形可用的数据逐个选择。

**字形轮廓变形（精确方式）。** `src/ui/src/pptx/render/text-warp-envelope-layout.ts` 将每个段落先按未变形状态排成一个整体块（`text-warp-envelope-block.ts`），再对整个块建立单一的连续映射（`text-warp-envelope-map.ts` 中的 `createEnvelopeWarp`，其 `P(s, v) = (1 - v) * T(s) + v * B(s)` 规律见该模块自身的文档注释）。能取得字形实际字体**文件**时，`text-warp-glyph-outline.ts` 使用 `opentype.js` 解析文件，并将字形真实矢量轮廓上的每个点，包括曲线上的点和曲线外的贝塞尔控制点，按**该点自身的位置**经过该映射，而非只处理字形边缘或中心。随后将字形渲染为单个变形后的 SVG `<path>`，不使用 `<text>` 元素。单元测试 `text-warp-glyph-outline.test.ts` 固定了以下不变量：“直线”包络会使每个点保持不变；`can` 预设的轮廓在整行中保持相同的垂直比例（圆柱上下曲线具有**相同**的半径和扫过角，仅相差固定偏移）。

实际字体字节可从两处取得（`text-warp-outline-font-cache.ts`）：文稿中**嵌入**的字体，其字节已由加载流程放入内存，同步解析，无需网络；以及已为屏幕文本解析出的 Google Fonts 目录网页字体，由 `text-warp-outline-webfont-fetch.ts` 在 CSS `<link>` 之外尽力单独获取实际 `.woff2` 字节。两者都得不到字体文件时，`text-warp-glyph-trace.ts` 会追踪浏览器自身渲染的字形轮廓（离屏 canvas 使用 `<text>` 回退所用的同一 CSS 字体栅格化字形，再对覆盖栅格按 50% 进行 marching-squares 等值线提取），因此在浏览器中，即使确实取不到字体**文件**，也不再回退到仿射近似：只有在没有 DOM canvas 的环境（例如服务端渲染）中，`createGlyphOutlineLookup` 才会使用仿射 / 切片仿射拟合（`text-warp-glyph-slicing.ts`）。这三条路径以相同方式接入五种绑定（React、Vue、Angular、Svelte、原生 JavaScript）。

**字形水平位置：弧长（2026-09-25 COM 重新测量）。** `createEnvelopeWarp` 按沿预设上下曲线的**弧长**比例确定字形的水平位置，因此曲线上陡峭的部分（`can` 半椭圆的垂直两端，强 `deflate` 的肩部）会压缩该处的字形，并将其余字形向外推。2026-09-24 的一次变更曾短暂改用简单线性放置（`x = s * width`），理由是 `textCanUp` 的 adj 15000 与 adj 50000 之间字干位置逐位相同；该证据无效，因为 `textCanUp` 将 `adj` 限制在 [66667, 100000]（`textCanDown` 限制在 [0, 33333]），两个值都会钳制为 66667，绘制出相同的形状。下方对整个轮廓的重新测量恢复了弧长方案：线性放置下，最深的 `can` 形状墨迹 IoU 只有 0.35 至 0.39，而使用弧长为 0.94 至 0.95。由于轮廓变形通过同一映射采样每个字形坐标，字形的宽度和剪切由曲线在其自身位置的局部斜率自然决定，无需为各预设做特殊处理。

**内部轮廓误差实测（2026-09-25）。** 方法：`scripts/make-wordart-can-fixture.ts` 每页写入一个黑色粗体艺术字标题（600×285pt 框，40pt，居中），文字为 `WORDART` 和 `Hamburgefonts`，字体为 Noto Sans、Verdana 和 Arial，预设为 `textCanUp` adj 66667 / 80000 / 92000 和 `textCanDown` adj 8000 / 20000 / 33333（均在限制范围内），另有 Arial 的 `textInflate` 10000 / 20000 和 `textDeflate` 18750 / 37500 对照，以及 8 根字干 `IIIIIIII` 标题的独立 20 页 `adj` 扫描。`scripts/measure-wordart-can-com.ps1` 通过 PowerPoint COM `Slide.Export` 以 1920×1080 导出每一页（每 pt 为 2.67px，因此框为 1600×760px）。`scripts/measure-wordart-can-viewer.mjs` 将同一文稿载入原生 JavaScript 演示（从源码解析 `ooxml-ui/pptx`），在设备缩放比 2 的无界面 Chromium 中截取主幻灯片舞台，将其精确的亚像素矩形映射到 1920×1080，并在亮度 128 处对两张图像做阈值化。每页的指标：墨迹 IoU（两个掩码的交集除以并集），以及对称轮廓距离（任一掩码的每个边界像素到另一个掩码最近边界像素的欧氏距离；平均值与第 95 百分位，单位为幻灯片宽 1920px 下的 px）。数值为两个字符串的平均值。

| 情况（1920px 宽）                              | IoU 线性 -> 弧长           | 平均 px 线性 -> 弧长   | p95 px 线性 -> 弧长    |
| ---------------------------------------------- | -------------------------- | ---------------------- | ---------------------- |
| `can` 最深（Up 66667 / Down 33333），3 种字体  | 0.35-0.39 -> 0.94-0.95     | 11.4-11.8 -> 1.1-1.2   | 29-30 -> 2.6-3.0       |
| `can` 较浅（Up 92000 / Down 8000），3 种字体   | 0.86-0.88 -> 0.95-0.97     | 2.9-3.4 -> 0.8-1.2     | 6.5 -> 1.5-2.5         |
| `can` 中间（Up 80000 / Down 20000），3 种字体  | 0.64-0.67 -> 0.75-0.77     | 8.5-9.0 -> 5.8         | 19 -> 17-18            |
| `textInflate` 10000 / 20000（Arial）           | 0.96 / 0.84 -> 0.97 / 0.97 | 1.0 / 3.7 -> 0.6 / 0.6 | 2.5 / 7.4 -> 1.5 / 1.5 |
| `textDeflate` 18750 / 37500（Arial）           | 0.86 / 0.57 -> 0.97 / 0.96 | 3.2 / 9.0 -> 0.6 / 0.7 | 6.5 / 19 -> 1.5 / 1.5  |
| 8 根字干扫描，20 个 `adj` 值中的 11 个（见下） | 0.59-0.97 -> 0.96-0.97     | 1.6-22.4 -> 1.4-1.7    | 4-48 -> 4-6            |

中间 `adj` 的几行受限于 PowerPoint 的某个伪影，而不是这条规律本身。字干扫描直接显示了这一点：在扫描的 20 个 `adj` 值中，有 9 个（`textCanUp` 80000 / 83333 / 86667 / 93333，`textCanDown` 3333 / 6667 / 13333 / 20000 / 23333），PowerPoint 会让文字在两条路径之一（上、下或两者）上提前结束，没有走到该路径的末端，因此该行标题在框右边缘之前约 20px（约为框宽的 1.2%）结束；当只有一行变短时，每个字形都会倾斜，右端最多约 25px。该效应对 `adj` 并不单调（80000 使下行变短，83333 使两行都变短，86667 使上行变短，90000 两行都不变短），这指向 PowerPoint 内部的曲线展平处理，而不是任何闭式规律，因此保持不建模；这 9 种情况在两种规律下的 IoU 都是 0.73 至 0.86。其余 11 个扫描值，即 PowerPoint 让两行都走到路径末端的情况，IoU 为 0.964 至 0.971，平均轮廓误差 1.4 至 1.7px，p95 为 4 至 6px，除 1px 外没有刚性偏移（最佳整数平移最多使 IoU 增加 0.02）。

多段落包络块的行不会再整体颠倒顺序：`layoutEnvelopeBlock` 将所有段落针对**同一个**块框排版，并且不会拉伸短行以填满框宽（短行保持自己的自然宽度，在块内居中或对齐），因此 `createEnvelopeWarp` 的单一连续、与行无关的映射会对每一行单调地插值 `v`（垂直比例），不再需要此前可能给极短且大幅拉伸的段落算出重叠带的逐行切带方式。`text-warp-envelope-layout.test.ts` 用一对对抗性的短段落与长段落固定了这一行为（在任何共同的水平位置都不会出现行交叉）；被它取代的 `edgeBandAt` / `text-warp-glyph-matrix.ts` 模块及其 58.7px / 48.8px 的部分修复测量已不再存在。

### 测量来源 {#measurement-provenance-1}

`textCanUp` / `textCanDown` 的圆柱上下 `arcTo` 曲线平行（半径和中心相同、相差常量偏移），因此每个字形保持相同高度，同时沿弧线剪切（已经 COM 验证）；`inflate` / `deflate` 及其变体则确实改变字形高度。`e2e/pptx/text-warp-envelope-parity.spec.ts` 中，单段落和 `wide-glyph-can` 切片情况的跨绑定偏差约在 1.5px 以内；上述多段落带顺序测试尚未通过。此前 `wide-glyph-can` 中原生 JavaScript 与其他绑定的切片数不一致（原生为 4 片，其他为 5 片）已修复：原生 JavaScript 的艺术字渲染器是唯一在测量字形推进宽度前跳过共享 PANOSE 字体替换步骤的绑定，导致字形切片采样不同的曲线位置。这是真实的绑定实现差异，不是测量竞争。

**已对照字体文件验证追踪轮廓（2026-09-25）。** 同一测量在相同字形上驱动两种轮廓来源：普通文稿（本地安装了 Noto Sans、Verdana 和 Arial，因此查看器不获取网页字体，而是从浏览器自身的渲染中追踪每个字形），以及由 PowerPoint 打开“嵌入字体”后重新保存的副本（`scripts/embed-wordart-can-fonts-com.ps1`；PowerPoint 嵌入了 Noto Sans 和 Verdana，但没有嵌入 Arial），此时查看器改为解析嵌入的字体文件。与 PowerPoint 对比，在每个 Noto Sans 和 Verdana 情况下，两条路径的得分相差都在 0.002 IoU 和 0.05px 平均轮廓距离以内（例如 Verdana `textCanUp` 66667：追踪为 0.952，字体文件为 0.954），因此 50% 阈值的 marching-squares 追踪在幻灯片尺度上没有可测的阈值或基线偏差。同一批运行中还有一点观察：在嵌入字体的文稿中，首个绘制的幻灯片保留了追踪轮廓（其路径数据与普通文稿完全一致），而之后每一页都使用了字体文件；两者吻合得非常接近，肉眼无法察觉，但这表明在首次渲染之后才完成的嵌入字体注册，不会让该幻灯片重新渲染。

## 相关阅读 {#related-reading}

- [已知限制](/zh/guide/limitations)：仍未解决的差距。
- [OpenXML 符合性](/zh/architecture/openxml-conformance)：包级覆盖清单，不代表视觉还原程度。
