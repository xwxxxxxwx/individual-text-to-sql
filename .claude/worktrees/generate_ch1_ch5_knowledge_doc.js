const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat,
  BorderStyle, WidthType, ShadingType,
  HeadingLevel, PageNumber
} = require('docx');

const FONT = 'Microsoft YaHei';
const CODE_FONT = 'Consolas';

function P(text, opts) {
  opts = opts || {};
  return new Paragraph({
    spacing: { after: 80, line: 320 },
    children: [new TextRun({ text: text, font: FONT, size: 22, ...(opts.run || {}) })]
  });
}
function H1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, pageBreakBefore: true,
    spacing: { before: 240, after: 180 },
    children: [new TextRun({ text: text, font: FONT, size: 36, bold: true, color: '1F4E79' })]
  });
}
function H2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text: text, font: FONT, size: 28, bold: true, color: '2E74B5' })]
  });
}
function H3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 160, after: 80 },
    children: [new TextRun({ text: text, font: FONT, size: 24, bold: true, color: '1F4E79' })]
  });
}
function Quote(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120, line: 320 },
    indent: { left: 360 },
    border: { left: { style: BorderStyle.SINGLE, size: 24, color: '5B9BD5', space: 8 } },
    children: [new TextRun({ text: text, font: FONT, size: 22, italics: true, color: '404040' })]
  });
}
function Warn(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120, line: 320 },
    indent: { left: 360 },
    border: { left: { style: BorderStyle.SINGLE, size: 24, color: 'E0883C', space: 8 } },
    shading: { fill: 'FDF2E3', type: ShadingType.CLEAR },
    children: [new TextRun({ text: text, font: FONT, size: 22, color: '8B4513' })]
  });
}
function MP(segments) {
  return new Paragraph({
    spacing: { after: 80, line: 320 },
    children: segments.map(s => new TextRun({
      text: s.text, font: s.code ? CODE_FONT : FONT,
      size: s.code ? 20 : 22, bold: s.bold || false, italics: s.italics || false,
      color: s.color || undefined,
      shading: s.code ? { fill: 'F4F4F4', type: ShadingType.CLEAR } : undefined,
    }))
  });
}
function bulletP(segments, level) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: level || 0 },
    spacing: { after: 60, line: 300 },
    children: segments.map(s => new TextRun({
      text: s.text, font: s.code ? CODE_FONT : FONT,
      size: s.code ? 20 : 22, bold: s.bold || false, italics: s.italics || false,
      color: s.color || undefined,
      shading: s.code ? { fill: 'F4F4F4', type: ShadingType.CLEAR } : undefined,
    }))
  });
}
function codeBlock(lines) {
  return lines.map(line => new Paragraph({
    spacing: { after: 0, line: 280 },
    shading: { fill: 'F4F4F4', type: ShadingType.CLEAR },
    indent: { left: 360 },
    children: [new TextRun({ text: line, font: CODE_FONT, size: 20 })]
  }));
}
const border = { style: BorderStyle.SINGLE, size: 6, color: 'BFBFBF' };
const allBorders = { top: border, bottom: border, left: border, right: border };
function cell(content, opts) {
  opts = opts || {};
  const isHeader = opts.header || false;
  return new TableCell({
    borders: allBorders,
    width: { size: opts.width, type: WidthType.DXA },
    shading: isHeader ? { fill: '5B9BD5', type: ShadingType.CLEAR } :
      (opts.shade ? { fill: opts.shade, type: ShadingType.CLEAR } : undefined),
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({
        text: String(content), font: opts.code ? CODE_FONT : FONT,
        size: opts.code ? 20 : 21, bold: isHeader || opts.bold || false,
        color: isHeader ? 'FFFFFF' : (opts.color || '000000')
      })]
    })]
  });
}
function makeTable(headers, rows, widths) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => cell(h, { header: true, width: widths[i], align: AlignmentType.CENTER }))
      }),
      ...rows.map(row => new TableRow({
        children: row.map((c, i) => {
          if (typeof c === 'object' && c !== null) return cell(c.text, Object.assign({}, c, { width: widths[i] }));
          return cell(c, { width: widths[i] });
        })
      }))
    ]
  });
}

const c = [];

// TITLE
c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 2400, after: 240 },
  children: [new TextRun({ text: 'STATS 762', font: FONT, size: 56, bold: true, color: '1F4E79' })] }));
c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 },
  children: [new TextRun({ text: '知识点详解', font: FONT, size: 48, bold: true, color: '1F4E79' })] }));
c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 480 },
  children: [new TextRun({ text: 'Ch1 – Ch5', font: FONT, size: 36, color: '2E74B5' })] }));
c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 },
  children: [new TextRun({ text: 'Regression for Data Science', font: FONT, size: 28, italics: true, color: '595959' })] }));
c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1200, after: 240 },
  children: [new TextRun({ text: '期末复习自学文档', font: FONT, size: 22, color: '808080' })] }));

// ============ Ch1 ============
c.push(H1('Ch1 · 线性回归模型基础'));
c.push(P('Essentials of Linear Regression Models', { run: { italics: true, color: '595959' } }));
c.push(Quote('一句话定位:把 Normal / Logistic / Poisson 回归用 GLM 三件套统一起来,并复习 ML 估计与统计推断。本章是后续所有章节的语言基础。'));

c.push(H2('1.1 课程定位'));
c.push(P('STATS 762 是 STATS 330 的研究生进阶版。前置课程:STATS 210(概率论 + 似然方法)、STATS 201/208(三种基础回归分别学过)。STATS 762 第一件事是用 GLM 把这三种回归统一进同一个框架。主要工具是 R + R Markdown。'));

c.push(H2('1.2 GLM 三件套(必背)'));
c.push(P('每一个 GLM 都由三个零件构成。这三件套是整门课的核心语言。'));
c.push(bulletP([{ text: '① 响应分布 ', bold: true }, { text: '(Random component):Y 服从指数族 (exponential family) 中的某个分布,如 Normal / Binomial / Poisson / Gamma / Negative Binomial。' }]));
c.push(bulletP([{ text: '② 线性预测器 ', bold: true }, { text: '(Systematic component):' }, { text: 'η = β0 + β1·x1 + ... + βk·xk', code: true }, { text: '。注意:即使加入 x² 或交互项,它对系数 β 仍然线性。' }]));
c.push(bulletP([{ text: '③ 链接函数 ', bold: true }, { text: '(Link function) ' }, { text: 'g(μ) = η', code: true }, { text: ':把均值 μ 的有界范围撑开到 (-∞, +∞),匹配线性预测器。' }]));

c.push(H2('1.3 三种典型 GLM 对照'));
c.push(makeTable(
  ['元素', 'Normal 回归', 'Logistic 回归', 'Poisson 回归'],
  [
    [{ text: '响应 Y', bold: true }, '连续', '二元成功数', '计数'],
    [{ text: '分布', bold: true }, 'N(μ, σ²)', 'Binomial(n, p)', 'Poisson(μ)'],
    [{ text: '均值', bold: true }, 'μ', 'n·p', 'μ'],
    [{ text: '方差', bold: true }, { text: 'σ² (独立)', color: 'C00000' }, { text: 'n·p·(1-p) (依赖 p)', color: 'C00000' }, { text: 'μ (= 均值)', color: 'C00000' }],
    [{ text: 'Canonical Link', bold: true }, 'identity', { text: 'logit: log(p/(1-p))', code: true }, { text: 'log(μ)', code: true }],
    [{ text: '典型例子', bold: true }, 'Catheter', 'CHD (年龄→心脏病)', 'Crab (体重→卫星数)']
  ],
  [1700, 2400, 2700, 2560]
));
c.push(P(''));
c.push(MP([{ text: '重要观察:', bold: true, color: 'C00000' }, { text: 'Normal 的方差 σ² 是独立参数,Logistic 和 Poisson 的方差由均值锁定 → 这就是 Ch2 过离散问题只发生在后两者的根源。' }]));

c.push(H2('1.4 链接函数的作用'));
c.push(P('链接函数把均值 μ 的有界范围映射到无界尺度,让线性预测器 (-∞, +∞) 能拟合。'));
c.push(P('以 Logistic 为例:直接套线性回归 P(CHD|age) = β0 + β1·age 会出问题,因为右边可以取任意值,但左边是概率必须在 [0,1]。链接函数 logit(p) = log(p/(1-p)) 把 [0,1] 撑开到 (-∞, +∞) 解决这个矛盾。'));
c.push(P('常见反链接函数(从 η 回到 μ/p):'));
c.push(...codeBlock([
  'Normal:    μ = η',
  'Logistic:  p = exp(η) / (1 + exp(η))    ← 也叫 expit / sigmoid',
  'Poisson:   μ = exp(η)',
  'Probit:    p = Φ(η)                      ← Φ 是标准正态 CDF'
]));

c.push(H2('1.5 Probit vs Logit(替代链接函数)'));
c.push(P('Probit 是 Logistic 的备选 link。两者都把 [0,1] 映射到 (-∞,+∞),但形态略有不同。'));
c.push(bulletP([{ text: '中心区域 (p 在 0.2~0.8):两者几乎重合,结果几乎一样' }]));
c.push(bulletP([{ text: '极端区域 (p 接近 0 或 1):probit 收敛更快(尾部更瘦)' }]));
c.push(bulletP([{ text: '实务选择:通常用 logit,因为系数有 log-odds 的直接解释' }]));
c.push(MP([{ text: 'R 代码:', bold: true }, { text: ' glm(..., family=binomial(link=PROBIT))', code: true }, { text: '  ← PROBIT 处替换为引号包围的 probit' }]));

c.push(H2('1.6 系数 β 的解释(必背)'));
c.push(P('β 永远直接作用在线性预测器 η 上,但翻译到 μ/p 的方式取决于 link function。'));
c.push(makeTable(
  ['模型', 'x+1 时 η 的变化', 'x+1 时 μ/p 的变化', '性质'],
  [
    [{ text: 'Normal', bold: true }, 'η +β', 'μ +β', '加性'],
    [{ text: 'Logistic', bold: true }, 'log(odds) +β', 'odds ×exp(β)', '乘性(odds)'],
    [{ text: 'Poisson', bold: true }, 'log(μ) +β', 'μ ×exp(β)', '乘性(均值)']
  ],
  [1500, 2300, 2700, 2860]
));
c.push(P(''));
c.push(MP([{ text: '示例:', bold: true }, { text: 'CHD 模型 logit(p) = -5.28 + 0.11·age。25 岁 vs 35 岁的 odds 之比 = exp(0.11 × 10) ≈ 3.0。注意:是 odds 乘 3 倍,不是 p 乘 3 倍!odds=1 对应 p=0.5,odds=3 对应 p=0.75。' }]));

c.push(H2('1.7 似然 (Likelihood) 的本质'));
c.push(P('似然 = 同一概率公式的反向解读:'));
c.push(bulletP([{ text: '概率', bold: true }, { text: ':已知参数 → 推数据可能性 (forward)' }]));
c.push(bulletP([{ text: '似然', bold: true }, { text: ':已知数据 → 把式子当作参数的函数,给每个候选参数打分 (backward)' }]));
c.push(P('最大似然估计 (MLE) 选打分最高的参数作为最优估计。这就是 GLM 求 β 的核心思想:哪个 β 最能解释观察到的数据,就选那个 β。'));
c.push(P('三种 GLM 的似然公式(结构都是每个观测的概率/密度连乘):'));
c.push(...codeBlock([
  'Normal:    L(β) = ∏ (2πσ²)^(-1/2) · exp(-(Y_i - μ_i)² / (2σ²))',
  'Logistic:  L(β) = ∏ C(n_i, Y_i) · p_i^Y_i · (1-p_i)^(n_i - Y_i)',
  'Poisson:   L(β) = ∏ μ_i^Y_i · exp(-μ_i) / Y_i!',
  '',
  '其中 μ_i / p_i 通过 link function 用 β 表达'
]));

c.push(H2('1.8 ML 估计的三步法'));
c.push(P('通用流程,不管哪种 GLM:'));
c.push(bulletP([{ text: 'Step 1:写出 log-likelihood ℓ(β) = log L(β)(连乘变连加,求导更方便)' }]));
c.push(bulletP([{ text: 'Step 2:对 β 求偏导 ∂ℓ/∂β' }]));
c.push(bulletP([{ text: 'Step 3:设偏导 = 0,解方程得到 β̂' }]));
c.push(P('等价表述(也常见):最大化 log L = 最小化 -log L = 最小化 Deviance(Ch2 会正式定义)。'));

c.push(H2('1.9 Normal 的解析解 + 其他 GLM 的 IRLS'));
c.push(P('Normal 回归是特例,第二步求出的方程是线性方程组,可以直接解:'));
c.push(...codeBlock([
  'β̂ = (X^T X)^(-1) X^T Y    ← Normal 解析解,也就是 OLS 公式'
]));
c.push(P('Normal 下 ML 估计 = OLS(最小二乘法)。两者给出同一个 β̂,因为 Normal 的 log-likelihood 等于 -RSS / (2σ²) + 常数,最大化 ℓ 等价于最小化 RSS。'));
c.push(P('Logistic / Poisson 的 ML 方程非线性,没有解析解,需要 IRLS(Iteratively Re-weighted Least Squares,反复加权最小二乘):'));
c.push(bulletP([{ text: '用当前 β 算 μ̂_i 和权重 w_i' }]));
c.push(bulletP([{ text: '做一次加权最小二乘得到新 β' }]));
c.push(bulletP([{ text: '重复直到 β 收敛' }]));
c.push(MP([{ text: '在 summary 输出里看:', bold: true }, { text: 'Number of Fisher Scoring iterations: 4', code: true }, { text: ' 就是 IRLS 收敛的迭代次数。' }]));

c.push(H2('1.10 推断的三大对象'));
c.push(P('拟合完模型后,推断 (inference) 关心三类问题:'));
c.push(makeTable(
  ['#', '推断对象', '工具', '在哪一章'],
  [
    ['1', '模型系数 β 是否显著', 'summary 里的 t/z value + p-value', 'Ch1'],
    ['2', '给定 X 下 Y 的预测区间', 'predict() + CI/PI', 'Ch1'],
    ['3', '哪些变量需要在模型里', 'F-test / Chi-squared / LRT / AIC', 'Ch2-Ch3']
  ],
  [600, 2900, 3500, 2360]
));

c.push(H2('1.11 t value vs z value 的区分(必考)'));
c.push(P('规则取决于 dispersion 参数 σ²/τ 是否需要从数据估计:'));
c.push(makeTable(
  ['模型类型', '用什么', '原因'],
  [
    [{ text: 'Normal', bold: true }, { text: 't value', color: 'C00000' }, 'σ² 估自数据,多一层不确定性'],
    [{ text: 'Logistic / Poisson 标准型', bold: true }, { text: 'z value', color: '2E7D32' }, 'dispersion 锁定 = 1,无额外不确定性'],
    [{ text: 'Quasi-Poisson / Quasi-Binomial', bold: true }, { text: 't value', color: 'C00000' }, 'τ 估自数据(Ch2)'],
    [{ text: 'Negative Binomial', bold: true }, { text: 'z value', color: '2E7D32' }, 'θ 通过迭代估,但用渐近正态推断']
  ],
  [3300, 1700, 4360]
));
c.push(P(''));
c.push(MP([{ text: '看 summary 里那行 ', bold: true }, { text: 'Dispersion parameter ... taken to be 1', code: true }, { text: ' → 用 z;要估计 dispersion → 用 t。' }]));

c.push(H2('1.12 Confidence Interval vs Prediction Interval(必考)'));
c.push(P('两个都是区间,但含义截然不同:'));
c.push(makeTable(
  ['区间', '在问什么', '方差来源', '宽度'],
  [
    [{ text: 'CI (confidence)', bold: true }, '这一类 X 的平均 Y 在哪?', 'Var(μ̂)', '较窄'],
    [{ text: 'PI (prediction)', bold: true }, '某一个具体 Y 会落在哪?', 'Var(μ̂) + σ²', '较宽']
  ],
  [2300, 3500, 2200, 1360]
));
c.push(P(''));
c.push(MP([{ text: 'Catheter 例(ht=44, wt=35):', bold: true }, { text: 'CI = (31.69, 40.98) 谈平均;PI = (26.61, 46.06) 谈个体。PI 比 CI 宽出来的部分就是个体偏离均值的随机波动 σ²。' }]));

c.push(H2('1.13 Logistic CI 的两种方法(必考)'));
c.push(P('要算 logistic 模型在给定 X 下 p 的 95% CI。'));
c.push(H3('方法 1(粗糙):直接在概率尺度做正态近似'));
c.push(...codeBlock([
  'predict(model, newdata, type=response, se.fit=TRUE)',
  '→ 用 estimate ± 1.96 × se.fit',
  '',
  '问题:当 p 靠近 0 或 1 时,CI 可能出现负值或 > 1'
]));
c.push(H3('方法 2(推荐):在 linear predictor 尺度做 CI 再变换'));
c.push(...codeBlock([
  'Step 1: predict(model, newdata, type=link, se.fit=TRUE)',
  'Step 2: 在 log-odds 尺度算 CI: η̂ ± 1.96 × se.fit → [L, U]',
  'Step 3: 用 expit 变换两端: p = exp(x)/(1+exp(x))',
  '        最终 CI = (expit(L), expit(U))',
  '',
  '保证 CI 严格在 [0, 1] 内'
]));
c.push(MP([{ text: '原理:', bold: true }, { text: '正态近似只在对称无界的尺度上合理。log-odds 是 (-∞, +∞) → 近似可靠;概率 [0,1] 有界 → 近似失效。这个套路适用所有 GLM:在 link scale 做 CI,变换回 response scale。' }]));

c.push(H2('1.14 GLS 估计量分布 + 方差传播'));
c.push(P('更一般的设置下(Σ 是误差协方差矩阵),β̂ 的分布推导:'));
c.push(...codeBlock([
  'β̂ = (X^T Σ^(-1) X)^(-1) X^T Σ^(-1) Y = AY     (GLS 估计量)',
  '',
  'Step 1 正态性: AY 是 Y 的线性变换, Y~N → β̂~N',
  'Step 2 期望: E[β̂] = AXβ = β  ✓ 无偏',
  'Step 3 方差: Var[β̂] = AΣA^T = (X^T Σ^(-1) X)^(-1)',
  '',
  '结论: β̂ ~ N(β, (X^T Σ^(-1) X)^(-1))'
]));
c.push(P('方差传播(复合预测的 CI):'));
c.push(...codeBlock([
  'Var[c^T β̂] = c^T Σ̂_β c     其中 Σ̂_β = vcov(model)',
  '',
  'c 向量构造: 截距永远 = 1, 激活的哑变量 = 1, 其余 = 0',
  '',
  '示例: 想算 wool=B, tension=M 这一组的 η̂ 的 CI:',
  '      c = [1, 1, 1, 0]^T        (4 个系数对应位置)',
  '      η̂ ± 1.96 × sqrt(c^T Σ̂_β c)',
  '      用 exp 变回 μ 尺度: [exp(L), exp(U)]'
]));

c.push(H2('1.15 三大案例精简'));
c.push(H3('Catheter(Normal 回归)'));
c.push(P('12 个小孩的身高 / 体重 / 心导管长度。模型 Y ~ N(μ, σ²),μ = β0 + β1·ht + β2·wt。link = identity,直接 lm() 或 glm(family=gaussian)。考点:CI vs PI 区分、系数加性解释。'));
c.push(H3('CHD(Logistic 回归)'));
c.push(P('100 个人,age 预测是否得 CHD。模型 logit(p) = β0 + β1·age。Grouped data 用 glm(cbind(yes, no) ~ age, family=binomial)。考点:logit vs probit 几乎重合、两种 CI 方法对比、odds ratio 计算。'));
c.push(H3('Crab(Poisson 回归)'));
c.push(P('173 只马蹄蟹,体重 + 颜色(4 levels)预测卫星数。模型 log(μ) = β0 + β1·weight + β2·C2 + β3·C3 + β4·C4。考点:factor() + contrasts() 自动建 indicator、参照组解读、过离散问题(τ̂ = 3.28 → Ch2 处理)。'));

c.push(H2('1.16 易错点(必背 5 条)'));
c.push(makeTable(
  ['#', '易错点', '正确做法'],
  [
    ['1', { text: '把 β 当成 X 加 1 时 Y 的变化', color: 'C00000' }, '只有 Normal 是加性;Logistic/Poisson 是乘性 exp(β)'],
    ['2', { text: 'CI 当 PI 用 / 反过来', color: 'C00000' }, 'CI 谈均值,PI 谈个体;预测某病人用 PI'],
    ['3', { text: '在概率尺度直接做 CI', color: 'C00000' }, '应在 link scale 做完再变换,避免 p < 0 或 > 1'],
    ['4', { text: 'Odds 倍数 = 概率倍数', color: 'C00000' }, '不对!odds=1 ↔ p=0.5; odds=3 ↔ p=0.75'],
    ['5', { text: 'Identity link 用于 Poisson', color: 'C00000' }, '可能 μ < 0,需 start=c(正数,0,...) 或改用 log link']
  ],
  [400, 3700, 4860]
));

// ============ Ch2 ============
c.push(H1('Ch2 · 模型诊断与拟合优度'));
c.push(P('Diagnostics & Goodness of Fit', { run: { italics: true, color: '595959' } }));
c.push(Quote('一句话定位:拟合完模型后两条诊断线 —— 诊断(point-level)找问题点,GoF/Deviance(model-level)给整体评分,最后用 dispersion 修正分布假设错配。'));

c.push(H2('2.1 诊断的两类问题'));
c.push(makeTable(
  ['类型', '关注什么', '找什么'],
  [
    [{ text: '数据问题', bold: true }, '哪些观测异常', 'outliers / 高杠杆点 HLP / 多重共线'],
    [{ text: '模型问题', bold: true }, '假设是否成立', '线性性 / 正态性 / 独立 / 同方差']
  ],
  [1700, 2800, 4860]
));
c.push(P('三类影响点:'));
c.push(bulletP([{ text: 'Outlier(异常值):Y 异常,残差大' }]));
c.push(bulletP([{ text: 'High Leverage Point(高杠杆点):X 异常,在 predictor space 偏远' }]));
c.push(bulletP([{ text: 'Influential point(强影响点):既是 outlier 又是 HLP,真正搅动模型' }]));

c.push(H2('2.2 Pearson 残差(核心)'));
c.push(P('Pearson 残差是 GLM 残差的标准化形式:'));
c.push(...codeBlock([
  'r_i = (Y_i - μ̂_i) / sqrt(Var(Y_i | X_i))',
  '',
  'Normal:    r_i = (Y_i - μ̂_i) / σ̂',
  'Binomial:  r_i = (Y_i - n_i·p̂_i) / sqrt(n_i·p̂_i·(1-p̂_i))',
  'Poisson:   r_i = (Y_i - μ̂_i) / sqrt(μ̂_i)'
]));
c.push(P('除以 sqrt(Var) 是为了标准化 —— 原始残差 Y - μ̂ 在不同观测点的方差不同(特别是 Poisson/Binomial),需要标准化才能可比。'));
c.push(MP([{ text: '解读:', bold: true }, { text: 'r_i > 0 模型低估;r_i < 0 模型高估。' }]));

c.push(H2('2.3 Cook distance + Leverage'));
c.push(P('Cook distance 衡量移除观测 i 后 β̂ 整体变化:'));
c.push(...codeBlock([
  'C_i ≈ (β̂ - β̂_{-i})^T I(β̂) (β̂ - β̂_{-i}) / (k+1)'
]));
c.push(P('Leverage 是 Hat Matrix 的对角元:'));
c.push(...codeBlock([
  'Normal:  H = X(X^T X)^(-1) X^T,  h_ii ∈ [0,1], 只依赖 X',
  'GLM:     H = W^(1/2) X (X^T W X)^(-1) X^T W^(1/2),  W 依赖 μ̂',
  '         → h_ii 依赖 X 和 Y',
  '',
  '幂等性: H² = H → Var[r_i] = σ²(1 - h_ii)'
]));
c.push(P('R 默认 plot(glm_fit) 输出 4 张诊断图:Residuals vs Fitted(线性性)、Q-Q plot(正态性)、Scale-Location(同方差)、Residuals vs Leverage(影响点)。'));

c.push(H2('2.4 Deviance(核心定义)'));
c.push(MP([{ text: '必背公式:', bold: true, color: 'C00000' }, { text: 'Dev = 2·log L_sat - 2·log L_mod', code: true }]));
c.push(P('L_sat = saturated model 似然(μ̂_i = y_i 完美拟合),L_mod = 你的模型似然。'));
c.push(bulletP([{ text: 'Dev ≥ 0,越小越好' }]));
c.push(bulletP([{ text: 'Normal 特例:Dev = RSS = Σ(Y - μ̂)²' }]));
c.push(bulletP([{ text: 'Null deviance:空模型(只截距)的 deviance' }]));
c.push(bulletP([{ text: 'Residual deviance:你的模型的 deviance' }]));
c.push(bulletP([{ text: '差值 Null − Residual = 协变量解释的 deviance' }]));

c.push(H2('2.5 三种 GoF 检验'));
c.push(makeTable(
  ['检验', '适用', '公式', 'R 代码'],
  [
    [
      { text: 'F-test', bold: true }, 'Normal / Quasi',
      { text: 'f = (ΔRSS/Δdf) / (RSS_F/(n-k-1))', code: true },
      { text: 'anova(fit1, fit2)', code: true }
    ],
    [
      { text: 'χ² test', bold: true }, 'Logistic / Poisson',
      { text: 'χ² = Σ Pearson 残差² ~ χ²_(n-k-1)', code: true },
      { text: 'anova(glm, test=Chisq)', code: true }
    ],
    [
      { text: 'LRT', bold: true }, '嵌套 GLM',
      { text: 'd_o = Dev_S - Dev_F ~ χ²_v', code: true },
      { text: 'anova(fit1, fit2, test=Chisq)', code: true }
    ],
    [
      { text: 'F-test (quasi)', bold: true }, 'Quasi-Poisson/Binomial',
      { text: 'F = (ΔDev/Δdf) / τ̂', code: true },
      { text: 'anova(fit1, fit2, test=F)', code: true }
    ]
  ],
  [1700, 2200, 3300, 2160]
));
c.push(P(''));
c.push(Warn('χ² 近似失效条件:大多数 μ̂_i < 5(稀疏数据)→ 改用 Parametric Bootstrap。'));

c.push(H2('2.6 ANOVA 顺序敏感性(必懂)'));
c.push(P('R 的 anova() 默认做顺序检验(Type I),变量加入顺序不同 → p 值不同。Crab 例子展示得很清楚:'));
c.push(...codeBlock([
  'anova(glm(sats ~ weight + colour))  → weight p < 2e-16, colour p = 0.028',
  'anova(glm(sats ~ colour + weight))  → colour p = 3e-5, weight p = 3e-14'
]));
c.push(P('原因:顺序检验测的是在已加入变量基础上,新变量是否显著。三种可能解读:'));
c.push(bulletP([{ text: '(1) weight 驱动 satellites,colour 仅与 weight 相关' }]));
c.push(bulletP([{ text: '(2) colour 驱动,weight 仅与 colour 相关' }]));
c.push(bulletP([{ text: '(3) 两者都贡献' }]));
c.push(P('数据本身无法区分,需要领域知识 → 这正是 Ch4 因果推断要解决的问题。'));

c.push(H2('2.7 Parametric Bootstrap GoF'));
c.push(P('当 χ² 近似失效(μ̂_i 太小)时用。5 步法:'));
c.push(...codeBlock([
  'Step 1: D_obs = deviance(your_model)              # 记录观测 deviance',
  'Step 2: Y_sim ~ assumed_distribution(μ̂_i)         # 在 H_0 下模拟',
  '        simulate(model)$sim_1',
  'Step 3: 在 Y_sim 上拟合相同结构模型 → D_sim',
  'Step 4: 重复 N 次(如 1000)→ 经验分布 emp_dist',
  'Step 5: p = #{D_sim > D_obs} / N                  # 经验 p-value'
]));
c.push(MP([{ text: '决策方向(注意):', bold: true, color: 'C00000' }, { text: 'H_0 = 模型拟合好。p < 0.05 → 拒绝 → 模型差。' }]));
c.push(MP([{ text: '为什么必须用 parametric?', bold: true }, { text: '非参数 bootstrap 重采样观测数据,无法在 H_0 假设的分布下生成。GoF 需要 H_0 下的模拟。' }]));

c.push(H2('2.8 过离散(Overdispersion)'));
c.push(P('理论方差:Poisson 假设 Var(Y) = μ,Binomial 假设 Var(Y) = n·p·(1-p)。现实可能 Var > 理论 → 过离散。'));
c.push(MP([{ text: '检测:', bold: true }, { text: 'τ̂ = Residual deviance / Residual df', code: true }]));
c.push(makeTable(
  ['τ̂ 值', '含义'],
  [
    ['τ̂ ≈ 1', '正常'],
    [{ text: 'τ̂ >> 1(如 > 1.5)', bold: true, color: 'C00000' }, '过离散'],
    ['τ̂ << 1', '欠离散(罕见)']
  ],
  [3000, 6360]
));
c.push(P(''));
c.push(MP([{ text: 'Crab 例:', bold: true }, { text: '551.78 / 168 = 3.28 → 严重过离散' }]));
c.push(P('两大成因:'));
c.push(bulletP([{ text: '遗漏变量 (unobserved heterogeneity):缺关键 X → 额外变异' }]));
c.push(bulletP([{ text: '相关观测:独立性假设被违反' }]));
c.push(P('影响:'));
c.push(bulletP([{ text: 'β̂ 仍然 unbiased(不受影响)' }]));
c.push(bulletP([{ text: 'SE 被低估 → p 值偏小 → 假阳性' }]));
c.push(bulletP([{ text: 'CI 太窄' }]));

c.push(H2('2.9 处理方案:Quasi-likelihood'));
c.push(...codeBlock([
  'Quasi-Poisson:   Var(Y) = τ · μ            # 引入自由参数 τ',
  'Quasi-Binomial:  Var(Y) = τ · n·p·(1-p)',
  '',
  'R: glm(..., family=quasipoisson)',
  '   glm(..., family=quasibinomial)'
]));
c.push(P('后果:'));
c.push(bulletP([{ text: 'β̂ 不变' }]));
c.push(bulletP([{ text: 'SE 膨胀为 sqrt(τ̂) 倍' }]));
c.push(bulletP([{ text: 'z-test 变成 t-test(τ 估自数据)' }]));
c.push(bulletP([{ text: '失去 likelihood → 不能用 LRT / AIC' }]));
c.push(bulletP([{ text: '模型比较必须用 F-test (test=F)' }]));
c.push(Warn('Quasi 模型不能用 test=Chisq!这是最常见的错误。'));

c.push(H2('2.10 处理方案:Negative Binomial'));
c.push(...codeBlock([
  'Y ~ NegBin(μ, θ)',
  'Var(Y) = μ + μ²/θ',
  '',
  'θ → ∞ 时退化为 Poisson',
  '',
  'R: library(MASS)',
  '   glm.nb(y ~ x1 + x2, data=...)'
]));
c.push(P('Quasi vs NB 对比:'));
c.push(makeTable(
  ['方面', 'Quasi-Poisson', 'Negative Binomial'],
  [
    ['似然', { text: '无', color: 'C00000' }, { text: '有', color: '2E7D32' }],
    ['LRT / AIC', { text: '✗', color: 'C00000' }, { text: '✓', color: '2E7D32' }],
    ['F-test', { text: '✓(必须)', color: '2E7D32' }, '✓'],
    ['处理 underdispersion', { text: '✓', color: '2E7D32' }, { text: '✗', color: 'C00000' }],
    ['方差结构', 'τ·μ(线性)', 'μ + μ²/θ(二次)'],
    ['何时优先', '简单 / 含欠离散', '需 likelihood 推断']
  ],
  [3000, 3000, 3360]
));

c.push(H2('2.11 Offsets'));
c.push(P('处理率(rate)数据:每 10 万人死亡数、每分钟纺织机故障数等。需要调整暴露量 (exposure) 让不同观测可比。'));
c.push(P('数学推导:'));
c.push(...codeBlock([
  '设 μ_x = 人口 x 下的均值,μ_100k = 标准人口下均值',
  '  μ_x = (x / 100k) × μ_100k',
  '取 log:  log(μ_x) = log(x/100k) + log(μ_100k)',
  '                    └────┬──────┘',
  '                    Offset 项,系数固定 = 1',
  '',
  'R: glm(deaths ~ Cytology + Age,',
  '       offset = log(pop/100000),',
  '       family = poisson)'
]));
c.push(MP([{ text: '必须用 log-link:', bold: true, color: 'C00000' }, { text: 'log-link 把乘性 (rate × exposure) 变成加性,offset 才能正确表达。identity link 下单位不匹配会失败。' }]));
c.push(MP([{ text: '系数解释变化:', bold: true }, { text: '从 expected count 变为 expected rate(每单位 exposure)。' }]));
c.push(P('答题模板:'));
c.push(Quote('Switching from wool A to B multiplies the expected rate of breaks (per minute of loom operation) by exp(β).'));

c.push(H2('2.12 易错点'));
c.push(makeTable(
  ['#', '易错点', '正确做法'],
  [
    ['1', { text: '混淆 Null 和 Residual deviance', color: 'C00000' }, 'Null = 空模型;Residual = 你的模型'],
    ['2', { text: 'τ̂ 方向搞反', color: 'C00000' }, 'τ̂ > 1 是过离散;τ̂ < 1 是欠离散'],
    ['3', { text: 'Quasi 用 Chisq test', color: 'C00000' }, 'Quasi 没有似然,必须用 F-test'],
    ['4', { text: '不看条件直接用 χ²', color: 'C00000' }, 'μ̂_i 大多 < 5 时 χ² 失效,改 parametric bootstrap'],
    ['5', { text: 'Offset 用 identity link', color: 'C00000' }, '必须 log-link'],
    ['6', { text: 'GoF p 值方向搞反', color: 'C00000' }, 'H_0 = 模型好,p < 0.05 → 拒绝 → 模型差']
  ],
  [400, 3700, 4860]
));

// ============ Ch3 ============
c.push(H1('Ch3 · 模型选择'));
c.push(P('Model Selection', { run: { italics: true, color: '595959' } }));
c.push(Quote('一句话定位:Ch2 的 F-test / LRT 一次只能比两个嵌套模型不够用;本章用信息准则 AIC / AICc / BIC 比较任意模型,并讨论搜索策略。'));

c.push(H2('3.1 模型选择的根本问题'));
c.push(P('p 个候选变量 → 2^p 个可能模型。Deviance test 局限:只能两两比较且必须嵌套 → 不够用。'));
c.push(P('信息论方法的优势:'));
c.push(bulletP([{ text: '可比较非嵌套模型' }]));
c.push(bulletP([{ text: '一次评估所有候选模型' }]));

c.push(H2('3.2 三大信息准则(必背公式)'));
c.push(makeTable(
  ['准则', '公式', '偏好', '何时用'],
  [
    [{ text: 'AIC', bold: true }, { text: '-2ℓ + 2k', code: true }, '较复杂', '大样本 + 关注预测损失'],
    [{ text: 'AICc', bold: true }, { text: 'AIC + 2k(k+1)/(n-k-1)', code: true }, '同 AIC', '小样本(n/k < 40)'],
    [{ text: 'BIC', bold: true }, { text: '-2ℓ + k·log(n)', code: true }, '较简单', '想找真模型']
  ],
  [1200, 3500, 1500, 3160]
));
c.push(P(''));
c.push(MP([{ text: 'AICc 修正目的:', bold: true }, { text: 'AIC 的惩罚项 2k 是大样本近似,n 小时低估复杂度惩罚 → AICc 补偿。n → ∞ 时修正项 → 0,AICc ≈ AIC。' }]));
c.push(P('AIC vs BIC 的根本区别:'));
c.push(bulletP([{ text: 'AIC:信息论视角,目标是最小化 KL 散度 → 偏向预测准确' }]));
c.push(bulletP([{ text: 'BIC:Bayesian 视角,目标是一致估计真模型 → 偏向模型简洁' }]));
c.push(Warn('警告 1:信息准则是相对指标,最小 AIC 不代表模型好,仍需诊断。警告 2:不能比较不同响应变量的模型(如 log(y) vs y)。'));

c.push(H2('3.3 GAM 用作非线性诊断'));
c.push(P('配合主模型外用 GAM 看每个 numeric covariate 的 partial effect 是否非线性。'));
c.push(P('判断方法:能否用一条直线穿过 dotted 置信带?'));
c.push(bulletP([{ text: '能 → 线性项足够' }]));
c.push(bulletP([{ text: '不能 → 考虑加 quadratic 项 / 更复杂形态' }]));
c.push(P('R:mgcv::gam()。Wine quality 例中发现 alcohol 和 sulphates 需要 quadratic 项。'));

c.push(H2('3.4 模型搜索策略'));
c.push(makeTable(
  ['策略', 'R 函数', '优缺点'],
  [
    [{ text: 'Exhaustive', bold: true }, { text: 'MuMIn::dredge()', code: true }, '拟合所有 2^p 模型,按 AICc 排序;变量少时优先'],
    [{ text: 'Backward stepwise', bold: true }, { text: 'step(full, direction=backward)', code: true }, '从满模型出发逐步删;适合大模型空间'],
    [{ text: 'Forward stepwise', bold: true }, { text: 'step(null, direction=forward)', code: true }, '从空模型出发逐步加'],
    [{ text: 'Both', bold: true }, { text: 'step(full, direction=both)', code: true }, '加减结合']
  ],
  [2400, 3700, 3260]
));
c.push(MP([{ text: '使用 dredge 前必须设置:', bold: true }, { text: 'options(na.action = na.fail)', code: true }]));
c.push(P('step() 的 k 参数:k=2 → AIC(默认);k=log(n) → BIC 惩罚。'));
c.push(MP([{ text: '贪心算法缺陷:', bold: true }, { text: 'Stepwise 可能错过全局最优。比如某两变量必须一起删才能降 AIC,stepwise 会漏掉。' }]));

c.push(H2('3.5 Numeric vs Factor 变量选择'));
c.push(makeTable(
  ['维度', 'Numeric 数值型', 'Factor 因子型'],
  [
    ['参数数量', '1 个系数', 'k-1 个哑变量'],
    ['假设形态', '线性 / 近似线性', '不假设形状'],
    ['外推能力', { text: '✓ 可以', color: '2E7D32' }, { text: '✗ 不能', color: 'C00000' }],
    ['适用场景', '关系单调 / 近似线性', '关系非单调 / 形状未知']
  ],
  [2400, 3300, 3660]
));
c.push(P('判断方法:'));
c.push(bulletP([{ text: '(1) 比较 AIC(小的好)' }]));
c.push(bulletP([{ text: '(2) 比较 Residual deviance / df(接近 1 好)' }]));
c.push(bulletP([{ text: '(3) LRT: d_o = Dev_num - Dev_fac ~ χ²_(df 差)' }]));
c.push(MP([{ text: '为什么是 k-1 个哑变量?', bold: true }, { text: 'k 个会完全多重共线(D_1 + D_2 + ... = 1);基准组被吸收入截距。' }]));

c.push(H2('3.6 易错点'));
c.push(makeTable(
  ['#', '易错点', '正确做法'],
  [
    ['1', { text: '用 AIC 比较 log(y) 和 y 的模型', color: 'C00000' }, '不能!响应变量不同的模型不可比'],
    ['2', { text: 'Stepwise 出的模型当成最优', color: 'C00000' }, '贪心算法可能错过全局最优,尽量用 dredge'],
    ['3', { text: 'AIC 最小就当作好模型', color: 'C00000' }, 'AIC 是相对指标,仍需做诊断'],
    ['4', { text: '嵌套模型只能用 F/LRT', color: 'C00000' }, '嵌套也可用 AIC/BIC;非嵌套只能 AIC/BIC']
  ],
  [400, 3700, 4860]
));

// ============ Ch4 ============
c.push(H1('Ch4 · 因果推断'));
c.push(P('Causal Models', { run: { italics: true, color: '595959' } }));
c.push(Quote('一句话定位:前三章在做 X 与 Y 关联吗,本章升级到 X 真的导致 Y 吗。用 counterfactual 定义 causation,用 DAG 画因果假设,再选合适变量估计直接 / 总因果效应。期末重点章。'));

c.push(H2('4.1 因果建模的两条路'));
c.push(bulletP([{ text: '设计实验 (RCT)', bold: true }, { text: ':gold standard,随机分配处理,随机化消除选择偏差' }]));
c.push(bulletP([{ text: '观测数据的因果分析', bold: true }, { text: ':本章重点,处理选择偏差需要 DAG + 变量控制' }]));

c.push(H2('4.2 因果性的数学定义'));
c.push(P('基于 counterfactual(反事实)的概念:'));
c.push(...codeBlock([
  'Y_i^{a=1} = 个体 i 在 A=1 时的输出',
  'Y_i^{a=0} = 个体 i 在 A=0 时的输出',
  '',
  '个体因果效应 (ICE): Y_i^{a=1} - Y_i^{a=0}',
  '  → 无法直接观察(同一时间只能经历一种处理)',
  '',
  '平均因果效应 (ACE): E[Y^{a=1}] - E[Y^{a=0}]',
  '  → 总体层面可估计'
]));
c.push(MP([{ text: 'RCT vs 观测研究:', bold: true }, { text: 'RCT 中组均值差 = ACE(随机化消除选择偏差);观测研究中组均值差 ≠ ACE(选择偏差未消除)。' }]));

c.push(H2('4.3 SUTVA 假设(两个条件)'));
c.push(P('Stable Unit Treatment Value Assumption,所有因果推断的基础假设:'));
c.push(bulletP([{ text: '条件 1 无干扰 (No Interference)', bold: true }, { text: ':个体的潜在结果只取决于自己的处理,不受他人处理影响。' }]));
c.push(bulletP([{ text: '条件 2 处理一致性 (Consistency)', bold: true }, { text: ':处理水平对所有单位含义相同,没有不同版本。' }]));
c.push(P('违反例子:'));
c.push(bulletP([{ text: '违反 ①:群体免疫(他人接种影响我的感染概率)' }]));
c.push(bulletP([{ text: '违反 ②:运动对不同人含义不同(散步 vs 高强度)' }]));

c.push(H2('4.4 Causal Diagram (DAG) 基础'));
c.push(P('DAG = Directed Acyclic Graph(有向无环图):'));
c.push(bulletP([{ text: '节点 = 变量' }]));
c.push(bulletP([{ text: '有向箭头 = 因果关系(A → B 表示 A 引起 B)' }]));
c.push(P('两个变量 A 和 B 出现相关的 5 种来源:'));
c.push(bulletP([{ text: '(1) 抽样误差(即使总体无关,样本里也有相关)' }]));
c.push(bulletP([{ text: '(2) A 引起 B(A → B)' }]));
c.push(bulletP([{ text: '(3) B 引起 A(A ← B)' }]));
c.push(bulletP([{ text: '(4) 共同原因 C 引起两者(A ← C → B)' }]));
c.push(bulletP([{ text: '(5) 选择偏差(A → S ← B,S 是选择条件)' }]));

c.push(H2('4.5 路径四分类(全章核心,必背)'));
c.push(P('从 S(变量)到 R(响应)的路径分四种:'));
c.push(makeTable(
  ['路径类型', '图示', '开/关', '估直接效应时', '估总效应时'],
  [
    [{ text: 'Direct(直接因)', bold: true }, 'S → R', '开', '保留', '保留'],
    [{ text: 'Indirect(间接因)', bold: true }, 'S → A → R', '开', { text: '关掉 A', color: 'C00000' }, '保留 A'],
    [{ text: 'Confounding(混杂)', bold: true }, 'S ← A → R', '开', { text: '关掉 A', color: 'C00000' }, { text: '关掉 A', color: 'C00000' }],
    [{ text: 'Colliding(碰撞)', bold: true }, 'S → A ← R', '天然关', { text: '不含 A', color: 'C00000' }, { text: '不含 A', color: 'C00000' }]
  ],
  [2400, 1700, 1100, 2080, 2080]
));
c.push(P(''));
c.push(MP([{ text: '操作核心:', bold: true, color: 'C00000' }, { text: '关闭路径 = 把变量加进模型当 regressor;混杂变量必加;Collider 绝不能加(加了反而打开虚假关联路径)。' }]));
c.push(P('Effect modifier(效应修饰因子):交互项,在 DAG 看不出来,需要 interaction term 测试。'));

c.push(H2('4.6 变量纳入规则'));
c.push(makeTable(
  ['变量类型', '结构', '估总效应', '估直接效应'],
  [
    ['Confounder 混杂', 'C → S 且 C → L', { text: '✓ 纳入', color: '2E7D32' }, { text: '✓ 纳入', color: '2E7D32' }],
    ['Mediator 中介', 'S → M → L', { text: '✗ 排除', color: 'C00000' }, { text: '✓ 纳入', color: '2E7D32' }],
    ['Collider 对撞', 'S → K ← L', { text: '✗ 永远排除', color: 'C00000' }, { text: '✗ 永远排除', color: 'C00000' }],
    ['其他直接影响 L 的', '→ L', { text: '✓ 纳入', color: '2E7D32' }, { text: '✓ 纳入', color: '2E7D32' }]
  ],
  [2400, 2000, 2400, 2560]
));
c.push(P(''));
c.push(MP([{ text: 'Mediation pattern:', bold: true }, { text: '总效应显著 + 直接效应不显著 → 效应完全通过中介变量传导。' }]));
c.push(MP([{ text: '重要警告:', bold: true, color: 'C00000' }, { text: '单个模型无法同时估计所有暴露变量的因果效应 —— 不同暴露的总/直接效应需要不同的协变量集合。' }]));

c.push(H2('4.7 简化决策模板'));
c.push(P('估直接效应 → 把所有直接影响 response 的变量全放进模型即可。'));
c.push(P('估总效应 → 只控制混杂变量;保留所有间接路径开放。'));
c.push(P('永远不要 → 加 collider(会引入虚假关联)。'));

c.push(H2('4.8 案例:College plans'));
c.push(P('4 变量 DAG:SES, IQ, PE, CP。SES + IQ 影响 PE;SES + IQ + PE 都直接影响 CP。'));
c.push(P('分析步骤:'));
c.push(bulletP([{ text: '估直接效应:用 PE + IQ + SES 都进模型的 logistic 回归' }]));
c.push(bulletP([{ text: '估 SES/IQ 对 CP 的总效应:模型只含 SES/IQ(不含 PE,让间接路径保持开放)' }]));
c.push(P('结果:总效应大于直接效应 → 间接效应(通过 PE 传导)与直接效应同向叠加。'));

c.push(H2('4.9 案例:Salary gap'));
c.push(P('UW 工资性别诉讼案。8 变量 DAG,gender 是关注变量。'));
c.push(P('直接效应模型:gender + 所有其他变量(field, rank, deg, startyr, yrdeg, admin) → β_gender 估计直接因果(控制所有中介后)。'));
c.push(P('总效应模型:只用 gender 作为唯一回归量 → β_gender 估计总因果(包括所有间接路径)。'));
c.push(P('结果对比:'));
c.push(bulletP([{ text: '直接效应:男性月薪比女性高约 $300(控制 rank、field 等之后仍然显著)' }]));
c.push(bulletP([{ text: '总效应:男性月薪比女性高约 $1100(包括性别影响 rank/field/admin 进而影响 salary 的间接通道)' }]));
c.push(MP([{ text: '解读:', bold: true }, { text: '总效应远大于直接效应 → 大部分性别差异通过间接路径(如晋升 rank 上)传导,而非直接歧视。' }]));

c.push(H2('4.10 漏变量的后果'));
c.push(makeTable(
  ['漏掉的类型', '直接效应估计', '总效应估计'],
  [
    ['漏混杂变量', { text: '有偏', color: 'C00000', bold: true }, { text: '有偏', color: 'C00000', bold: true }],
    ['漏 collider', { text: '不漏 OK(本来就不该加)', color: '2E7D32' }, { text: '同上', color: '2E7D32' }],
    ['漏间接路径中介', { text: '有偏', color: 'C00000', bold: true }, { text: '无偏', color: '2E7D32', bold: true }],
    ['漏单纯直接效应变量', '无偏但精度低', '同上']
  ],
  [3000, 3180, 3180]
));

c.push(H2('4.11 易错点'));
c.push(makeTable(
  ['#', '易错点', '正确做法'],
  [
    ['1', { text: '估总效应时加中介变量', color: 'C00000' }, '估总效应要保留间接路径开放,不加中介'],
    ['2', { text: '为了控制变量加 collider', color: 'C00000' }, 'Collider 绝不能加,反而会引入虚假关联'],
    ['3', { text: '用同一个模型估所有变量的因果效应', color: 'C00000' }, '不同变量需要不同的协变量集合'],
    ['4', { text: 'DAG 看不出 effect modifier 就忽略交互', color: 'C00000' }, '效应修饰要用 interaction term 单独测'],
    ['5', { text: 'SUTVA 当成处理不会改变', color: 'C00000' }, 'SUTVA = 无干扰 + 处理一致性,与变化无关']
  ],
  [400, 3700, 4860]
));

// ============ Ch5 ============
c.push(H1('Ch5 · 预测与交叉验证'));
c.push(P('Prediction, CV, Bootstrap', { run: { italics: true, color: '595959' } }));
c.push(Quote('一句话定位:前面章节关注拟合优度,本章把目光转向预测。用 K-fold CV 客观估计模型在新数据上的表现,用 Bootstrap 重采样得到任意统计量的经验抽样分布。'));

c.push(H2('5.1 预测 vs 推断'));
c.push(makeTable(
  ['', '统计推断 Inference', '预测 Prediction'],
  [
    ['目标', '理解系数关系', '最小化新数据误差'],
    ['关注', '显著性 / 因果', '预测误差大小'],
    ['工具', '假设检验 / CI', '交叉验证 / CV']
  ],
  [1500, 3700, 3160]
));

c.push(H2('5.2 Loss Function 与 Risk'));
c.push(P('Loss function 量化模型预测与真实值的差距。'));
c.push(...codeBlock([
  '连续 Y:',
  '  Squared error:  L = (y - ŷ)²            ← 大误差被重罚',
  '  Absolute error: L = |y - ŷ|',
  '  Log-likelihood: L = -2 log p_θ(Y)',
  '',
  '分类 Y:',
  '  0-1 loss:       L = 1(ŷ ≠ y)            ← 分类误差率',
  '  Log-likelihood: L = -2 Σ 1(Y=j) log p(ŷ=j)'
]));
c.push(MP([{ text: 'Risk', bold: true }, { text: ' = E[Loss],对未来数据的期望损失,无法直接计算 → 用 CV 估计。' }]));
c.push(MP([{ text: 'MSPE', bold: true }, { text: ' = 新数据上 (predicted - actual)² 的均值(非训练数据!)。' }]));

c.push(H2('5.3 K-Fold 交叉验证'));
c.push(P('算法:'));
c.push(bulletP([{ text: '数据随机分 K 份' }]));
c.push(bulletP([{ text: '轮流用 1 份测试,K-1 份训练' }]));
c.push(bulletP([{ text: '重复 K 次,CV 误差 R̂_K = K 次测试误差的平均' }]));
c.push(MP([{ text: '实践选择:', bold: true }, { text: 'K = 5 或 K = 10(偏差方差权衡的经验最优)。' }]));
c.push(P('数学表达:'));
c.push(...codeBlock([
  'Step 1: 数据分 K 折 S_1, ..., S_K',
  '',
  'Step 2: 对每个 k = 1, ..., K:',
  '  (i) 用 ∪_{j≠k} S_j 训练模型,得 θ_k',
  '  (ii) 在 S_k 上算误差: e_k = (1/n_k) Σ L(y_i, f_{θ_k}(x_i))',
  '',
  'Step 3: R̂_K = (1/K) Σ e_k'
]));

c.push(H2('5.4 LOOCV vs K-Fold'));
c.push(P('LOOCV = Leave One Out CV,即 K = n(每次留一个点测试)。'));
c.push(makeTable(
  ['', 'K-fold (K=5/10)', 'LOOCV (K=n)'],
  [
    ['偏差 Bias', '适中', { text: '最低(训练集 n-1)', color: '2E7D32' }],
    ['方差 Variance', { text: '低(折间相关 ρ 小)', color: '2E7D32' }, { text: '高(ρ ≈ 1,折间高度重叠)', color: 'C00000' }],
    ['计算量', { text: '快', color: '2E7D32' }, { text: '慢(需拟合 n 次)', color: 'C00000' }],
    ['适用', '默认首选', '数据量极小时']
  ],
  [2000, 3500, 3860]
));
c.push(P(''));
c.push(MP([{ text: '方差公式:', bold: true }, { text: 'Var(R̂_K) ≈ σ²(1/K + (K-1)/K × ρ)', code: true }, { text: '。ρ → 1 时无论 K 多大方差都高。' }]));
c.push(Warn('CV 误差没有通用无偏方差估计(折间相关)→ 当两个模型 CV 误差很接近时,无法可靠地比较选择。'));

c.push(H2('5.5 CV 的渐近性质'));
c.push(P('一致性:固定 K,n → ∞ 时 R̂_K - E[R(θ̄)] → 0(概率收敛到真风险)。'));
c.push(P('Bias 大小:Bias = O(1/(nK))。K 越大,bias 越小(LOOCV 最优)。'));
c.push(P('Variance 大小:和折间相关 ρ 有关。LOOCV 的 ρ ≈ 1 → variance 最大。'));
c.push(MP([{ text: 'Bias-variance trade-off(必背):', bold: true, color: 'C00000' }]));
c.push(makeTable(
  ['K', 'Bias', 'Variance'],
  [
    ['Small K', { text: 'High', color: 'C00000' }, { text: 'Low', color: '2E7D32' }],
    ['Large K (LOOCV)', { text: 'Low', color: '2E7D32' }, { text: 'High', color: 'C00000' }]
  ],
  [3000, 3000, 3360]
));

c.push(H2('5.6 Non-Parametric Bootstrap(非参数自助法)'));
c.push(P('解决问题:在某些情况下,理论抽样分布难以推导(如复杂统计量、小样本、分布假设不成立)。Bootstrap 通过重采样数据来经验估计抽样分布。'));
c.push(P('算法:'));
c.push(...codeBlock([
  '原始数据 (Z_1, ..., Z_n),统计量 S(z),估计量 Ŝ_n',
  '',
  'for b = 1, ..., B:',
  '  Step 1: 有放回采样得 Z*^b,大小与原始数据相同',
  '  Step 2: 在 Z*^b 上算 S(Z*^b)',
  '',
  '得到 B 个 S(Z*^b),构成经验抽样分布'
]));

c.push(H2('5.7 Bootstrap 估计量'));
c.push(...codeBlock([
  '样本均值: S̄* = (1/B) Σ S(Z*^b)',
  '',
  'Bias: Bias = S̄* - Ŝ_n',
  '',
  'Variance: Var(S*) = (1/B) Σ (S(Z*^b) - S̄*)²',
  '',
  '1-α 置信区间:',
  '  排序: S(Z*^1) < ... < S(Z*^B)',
  '  α/2 = #{S(Z*) < S_lower} / B',
  '  1-α/2 = #{S(Z*) < S_upper} / B'
]));

c.push(H2('5.8 Bootstrap 性质与注意事项'));
c.push(bulletP([{ text: 'S(Z*) - Ŝ_n 的分布 ≈ Ŝ_n - S(Z) 的分布' }]));
c.push(bulletP([{ text: 'Var(S*) → Var(Ŝ_n)(大数定律)' }]));
c.push(bulletP([{ text: '一致性:P(a_n(Ŝ_n - S) ≤ x) - P(a_n(S* - Ŝ_n) ≤ x) → 0' }]));
c.push(Warn('Bootstrap 不是用来得到更好的点估计 —— 重采样的期望就是原始样本估计。数据必须独立且代表总体。不能修正模型 misspecification。'));

c.push(H2('5.9 易错点'));
c.push(makeTable(
  ['#', '易错点', '正确做法'],
  [
    ['1', { text: 'LOOCV 一定比 K-fold 准', color: 'C00000' }, 'LOOCV bias 小但 variance 高,K=5/10 才是经验最优'],
    ['2', { text: '认为 bootstrap 能改进点估计', color: 'C00000' }, 'Bootstrap 估计的是 SE / CI / bias,不改进点估计'],
    ['3', { text: 'Bootstrap 修正 misspecification', color: 'C00000' }, 'Bootstrap 不能修正模型形态错误'],
    ['4', { text: 'Bootstrap CI 直接用 ±1.96·se', color: 'C00000' }, '用 percentile 方法(排序后取分位数)更稳'],
    ['5', { text: 'CV 误差当成最终评估', color: 'C00000' }, '需配合独立测试集才算最终评估']
  ],
  [400, 3700, 4860]
));

// ============ APPENDIX ============
c.push(H1('附录 A · 关键公式速查'));

c.push(H2('A.1 GLM 系数解释'));
c.push(...codeBlock([
  'Normal:    x+1 → μ 加 β        (加性)',
  'Logistic:  x+1 → log(odds) 加 β; odds × exp(β)',
  'Poisson:   x+1 → log(μ) 加 β; μ × exp(β)'
]));

c.push(H2('A.2 GLM 估计量分布'));
c.push(...codeBlock([
  'β̂ ~ N(β, (X^T Σ^(-1) X)^(-1))      ← GLS 一般形式',
  'Var[c^T β̂] = c^T Σ̂_β c              ← Σ̂_β = vcov(model)',
  '',
  'CI: β̂ ± 1.96 × SE                  ← 系数 95% CI',
  'CI: exp(η̂ ± 1.96 × sqrt(c^T Σ̂ c))  ← μ̂ 的 95% CI (log-link)'
]));

c.push(H2('A.3 残差与杠杆'));
c.push(...codeBlock([
  'Pearson 残差: r_i = (Y_i - μ̂_i) / sqrt(Var(Y_i|X_i))',
  'Hat matrix: H = X(X^T X)^(-1) X^T (Normal)',
  '            H = W^(1/2) X (X^T W X)^(-1) X^T W^(1/2) (GLM)',
  'Var[r_i] = σ²(1 - h_ii)'
]));

c.push(H2('A.4 拟合优度统计量'));
c.push(...codeBlock([
  'Deviance: Dev = 2(log L_sat - log L_fitted)',
  '  ≥ 0, 越小越好;Normal: Dev = RSS',
  '',
  'τ̂ = Resid.Dev / df    ≈ 1 正常; >> 1 过离散; << 1 欠离散',
  '',
  'F-test (Normal): F = [(RSS_S - RSS_F)/Δdf] / [RSS_F/(n-k-1)]',
  'F-test (Quasi): F = (ΔDev/Δdf) / τ̂',
  'LRT: d_o = Dev_sub - Dev_full ~ χ²_v',
  '',
  'Parametric Bootstrap p: p = #{D_sim > D_obs} / Nsim'
]));

c.push(H2('A.5 模型选择'));
c.push(...codeBlock([
  'AIC  = -2ℓ + 2k',
  'AICc = AIC + 2k(k+1)/(n-k-1)',
  'BIC  = -2ℓ + k·log(n)',
  '',
  '选择规则: 小样本用 AICc; 想找真模型用 BIC; 预测优先用 AIC'
]));

c.push(H2('A.6 交叉验证 + Bootstrap'));
c.push(...codeBlock([
  'K-fold: R̂_K = (1/K) Σ e_k',
  'Var(R̂_K) ≈ σ²(1/K + (K-1)/K × ρ)',
  '  Small K: 高 bias, 低 variance',
  '  LOOCV (K=n): 低 bias, 高 variance',
  '',
  'Bootstrap:',
  '  Bias = S̄* - Ŝ_n',
  '  Var(S*) = (1/B) Σ (S(Z*^b) - S̄*)²',
  '  CI: 排序 S(Z*^b),取 α/2 和 1-α/2 分位数'
]));

c.push(H1('附录 B · R 函数速查'));

c.push(H2('B.1 拟合'));
c.push(...codeBlock([
  'lm(y ~ x, data)                                    # Normal 等价',
  'glm(y ~ x, family=gaussian, data)                  # Normal',
  'glm(cbind(yes,no) ~ x, family=binomial, data)      # Logistic (grouped)',
  'glm(y ~ x, family=binomial(link=probit), data)     # Probit',
  'glm(y ~ x, family=poisson, data)                   # Poisson',
  'glm(y ~ x, family=quasipoisson, data)              # Quasi-Poisson',
  'MASS::glm.nb(y ~ x, data)                          # Negative Binomial',
  'glm(y ~ x + offset(log(exp)), family=poisson)      # Poisson + offset'
]));

c.push(H2('B.2 推断与预测'));
c.push(...codeBlock([
  'summary(fit)                                       # 系数表、deviance、AIC',
  'anova(fit1, fit2, test=Chisq or F)                 # 嵌套模型比较',
  'predict(fit, newdata, type=link or response)       # 预测',
  'predict(fit, newdata, interval=confidence/prediction)  # CI/PI',
  'vcov(fit)                                          # β̂ 的方差矩阵',
  'plot(fit)                                          # 4 张诊断图'
]));

c.push(H2('B.3 模型选择'));
c.push(...codeBlock([
  'AIC(fit), AICc(fit), BIC(fit)                      # 各信息准则',
  'MuMIn::dredge(full_fit)                            # 穷举搜索',
  'step(full_fit, direction=backward/forward/both)    # 贪心搜索',
  'mgcv::gam(y ~ s(x), data)                          # 非线性诊断',
  'factor(x); contrasts(x)                            # 哑变量',
  'simulate(fit)$sim_1                                # Parametric Bootstrap'
]));

c.push(H1('附录 C · 考试答题模板'));

c.push(H2('C.1 系数解释模板'));
c.push(Quote('Log-link (Poisson): For each one-unit increase in x, the expected count/rate is multiplied by exp(β) ≈ ___, holding others constant.'));
c.push(Quote('Logit (Logistic): For each one-unit increase in x, the log-odds increases by β; the odds is multiplied by exp(β) = ___.'));
c.push(Quote('Offset (rate model): Switching from A to B multiplies the expected rate of [outcome] by exp(β), per unit of [exposure].'));

c.push(H2('C.2 Overdispersion 答题模板'));
c.push(Quote('Identification: τ̂ = Resid.Dev / df = ___ >> 1, indicating overdispersion.'));
c.push(Quote('Consequence: SE is underestimated → p-values too small → false positives.'));
c.push(Quote('Remedy: Refit with quasipoisson; SE inflated by sqrt(τ̂); use F-test instead of LRT.'));

c.push(H2('C.3 Bootstrap GoF 答题模板'));
c.push(Quote('H_0: The model is correctly specified.'));
c.push(Quote('Conclusion: p = ___ < 0.05 → reject H_0 → model fits poorly.'));
c.push(Quote('Justification for using bootstrap (not chi-squared): Some μ̂_i < 5 → χ² approximation invalid → use parametric bootstrap.'));

c.push(H2('C.4 因果 DAG 分析模板'));
c.push(Quote('总效应: Include confounders; EXCLUDE mediators (keep indirect paths open).'));
c.push(Quote('直接效应: Include confounders AND mediators (close indirect paths).'));
c.push(Quote('永远不要: Include collider → opens spurious path.'));

c.push(H2('C.5 假设检验类型对照'));
c.push(makeTable(
  ['类型', 'H_0', '工具'],
  [
    ['单个系数 (Wald/z-test)', 'β_j = 0(变量无效)', 'summary() 里的 p 值'],
    ['嵌套模型 LRT', '新增系数全为零', 'anova(m1, m2, test=Chisq)'],
    ['嵌套模型 F-test (Quasi)', '新增系数全为零', 'anova(m1, m2, test=F)'],
    ['GoF', '模型拟合良好', 'anova(fit, test=Chisq) 或 simulate()'],
    ['Overdispersion', 'τ = 1(无过离散)', 'τ̂ = Resid.Dev/df'],
    ['因果总效应', '暴露变量无总效应', 'anova(reduced_fit, test=Chisq)']
  ],
  [2700, 2700, 3960]
));
c.push(P(''));
c.push(MP([{ text: '决策规则:', bold: true, color: 'C00000' }, { text: 'p < 0.05 → 拒绝 H_0;GoF 拒绝 H_0 意味着模型不好(与单变量检验方向反)。' }]));

c.push(H2('C.6 英文词汇表'));
c.push(makeTable(
  ['中文', '英文'],
  [
    ['统计显著', 'statistically significant'],
    ['表明 / 指示', 'indicate'],
    ['期望值', 'expected number / expected count'],
    ['对数尺度 / 均值尺度', 'log scale / mean scale'],
    ['参照基准', 'reference level'],
    ['被吸收入截距', 'absorbed into the intercept'],
    ['SE 被膨胀 sqrt(τ)', 'SE is inflated by sqrt(tau)'],
    ['p 值更保守', 'p-values become more conservative'],
    ['假阳性', 'false positive'],
    ['零假设', 'null hypothesis'],
    ['哑变量 / 设计向量', 'dummy values / design vector c'],
    ['方差矩阵', 'variance matrix = vcov(m)'],
    ['对角线 / 非对角线', 'diagonal / off-diagonal'],
    ['模拟的', 'simulated'],
    ['经验 p 值', 'empirical p-value'],
    ['模型应被修正', 'the model should be revised']
  ],
  [3500, 5860]
));

const doc = new Document({
  creator: 'STATS 762 Review',
  title: 'STATS 762 Knowledge Doc Ch1-Ch5',
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 36, bold: true, font: FONT, color: '1F4E79' },
        paragraph: { spacing: { before: 240, after: 180 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: FONT, color: '2E74B5' },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: FONT, color: '1F4E79' },
        paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } }
    ]
  },
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        { level: 1, format: LevelFormat.BULLET, text: '◦', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1440, hanging: 360 } } } }
      ]
    }]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: 'STATS 762 · 知识点详解 Ch1-Ch5', font: FONT, size: 18, color: '808080' })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: '第 ', font: FONT, size: 18, color: '808080' }),
            new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: '808080' }),
            new TextRun({ text: ' 页 / 共 ', font: FONT, size: 18, color: '808080' }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: '808080' }),
            new TextRun({ text: ' 页', font: FONT, size: 18, color: '808080' })
          ]
        })]
      })
    },
    children: c
  }]
});

const outputPath = 'C:\\Users\\33238\\Desktop\\2026 semester1 选课\\2.STATS 762\\STATS762_知识点详解_Ch1-Ch5.docx';
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log('Generated:', outputPath);
  console.log('Size:', (buffer.length / 1024).toFixed(1), 'KB');
});
