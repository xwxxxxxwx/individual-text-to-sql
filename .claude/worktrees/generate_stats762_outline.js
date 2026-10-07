const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, PageOrientation, LevelFormat,
  TabStopType, TabStopPosition, BorderStyle, WidthType, ShadingType,
  HeadingLevel, PageNumber, PageBreak
} = require('docx');

// --- helpers ---
const FONT = "Microsoft YaHei"; // Chinese-friendly
const CODE_FONT = "Consolas";

function P(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 80, line: 320 },
    ...opts,
    children: [new TextRun({ text, font: FONT, size: 22, ...opts.run })]
  });
}

function H1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    pageBreakBefore: true,
    spacing: { before: 240, after: 180 },
    children: [new TextRun({ text, font: FONT, size: 36, bold: true, color: "1F4E79" })]
  });
}

function H2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, font: FONT, size: 28, bold: true, color: "2E74B5" })]
  });
}

function H3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 160, after: 80 },
    children: [new TextRun({ text, font: FONT, size: 24, bold: true, color: "1F4E79" })]
  });
}

function Quote(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120, line: 320 },
    indent: { left: 360 },
    border: {
      left: { style: BorderStyle.SINGLE, size: 24, color: "5B9BD5", space: 8 }
    },
    children: [new TextRun({ text, font: FONT, size: 22, italics: true, color: "404040" })]
  });
}

function Warn(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120, line: 320 },
    indent: { left: 360 },
    border: {
      left: { style: BorderStyle.SINGLE, size: 24, color: "E0883C", space: 8 }
    },
    shading: { fill: "FDF2E3", type: ShadingType.CLEAR },
    children: [new TextRun({ text, font: FONT, size: 22, color: "8B4513" })]
  });
}

// Mixed-styled paragraph from segments: { text, bold?, code?, color? }
function MP(segments, opts = {}) {
  return new Paragraph({
    spacing: { after: 80, line: 320 },
    ...opts,
    children: segments.map(s => new TextRun({
      text: s.text,
      font: s.code ? CODE_FONT : FONT,
      size: s.code ? 20 : 22,
      bold: s.bold || false,
      italics: s.italics || false,
      color: s.color || undefined,
      shading: s.code ? { fill: "F4F4F4", type: ShadingType.CLEAR } : undefined,
    }))
  });
}

function bulletP(segments, opts = {}) {
  return new Paragraph({
    numbering: { reference: "bullets", level: opts.level || 0 },
    spacing: { after: 60, line: 300 },
    children: segments.map(s => new TextRun({
      text: s.text,
      font: s.code ? CODE_FONT : FONT,
      size: s.code ? 20 : 22,
      bold: s.bold || false,
      italics: s.italics || false,
      color: s.color || undefined,
      shading: s.code ? { fill: "F4F4F4", type: ShadingType.CLEAR } : undefined,
    }))
  });
}

function codeBlock(lines) {
  return lines.map(line => new Paragraph({
    spacing: { after: 0, line: 280 },
    shading: { fill: "F4F4F4", type: ShadingType.CLEAR },
    indent: { left: 360 },
    children: [new TextRun({ text: line, font: CODE_FONT, size: 20 })]
  }));
}

// --- Table helpers ---
const border = { style: BorderStyle.SINGLE, size: 6, color: "BFBFBF" };
const allBorders = { top: border, bottom: border, left: border, right: border };

function cell(text, opts = {}) {
  const isHeader = opts.header || false;
  return new TableCell({
    borders: allBorders,
    width: { size: opts.width, type: WidthType.DXA },
    shading: isHeader
      ? { fill: "5B9BD5", type: ShadingType.CLEAR }
      : (opts.shade ? { fill: opts.shade, type: ShadingType.CLEAR } : undefined),
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({
        text: String(text),
        font: opts.code ? CODE_FONT : FONT,
        size: opts.code ? 20 : 21,
        bold: isHeader || opts.bold || false,
        color: isHeader ? "FFFFFF" : (opts.color || "000000")
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
          // c can be string or {text, ...opts}
          if (typeof c === 'object' && c !== null) {
            return cell(c.text, { width: widths[i], ...c });
          }
          return cell(c, { width: widths[i] });
        })
      }))
    ]
  });
}

// ============ CONTENT ============
const content = [];

// Title page
content.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 2400, after: 240 },
  children: [new TextRun({ text: "STATS 762", font: FONT, size: 56, bold: true, color: "1F4E79" })]
}));
content.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 240 },
  children: [new TextRun({ text: "期末复习大纲", font: FONT, size: 48, bold: true, color: "1F4E79" })]
}));
content.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 480 },
  children: [new TextRun({ text: "期中前内容 (Ch1 – Ch5)", font: FONT, size: 36, color: "2E74B5" })]
}));
content.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 240 },
  children: [new TextRun({ text: "Regression for Data Science", font: FONT, size: 28, italics: true, color: "595959" })]
}));
content.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 1200, after: 240 },
  children: [new TextRun({ text: "基于课堂字幕与课件 PDF 整理", font: FONT, size: 22, color: "808080" })]
}));

// ===== Overall structure =====
content.push(new Paragraph({
  heading: HeadingLevel.HEADING_1,
  pageBreakBefore: true,
  spacing: { before: 240, after: 180 },
  children: [new TextRun({ text: "整体结构与依赖关系", font: FONT, size: 36, bold: true, color: "1F4E79" })]
}));

content.push(...codeBlock([
  "Ch1 (GLM 基础语言)",
  "  ├── Ch2 (诊断 + GoF + Deviance) ─────┐",
  "  │     └── Ch3 (基于 deviance 比较模型) │",
  "  │                                     ├── Ch5 (CV/Bootstrap, 视角转向预测)",
  "  └── Ch4 (因果推断 — 独立分支但思想贯穿) ─┘"
]));

content.push(P(""));
content.push(bulletP([{ text: "Ch1 ", bold: true }, { text: "给所有人一套统一语言（GLM = 分布 + 线性预测 + link）" }]));
content.push(bulletP([{ text: "Ch2 ", bold: true }, { text: "用 deviance 评估单个模型 + 修正分布假设" }]));
content.push(bulletP([{ text: "Ch3 ", bold: true }, { text: "把 deviance 升级成 AIC/BIC，扩展到任意模型比较" }]));
content.push(bulletP([{ text: "Ch4 ", bold: true }, { text: "平行于 Ch2-3,处理“X 是否导致 Y”" }]));
content.push(bulletP([{ text: "Ch5 ", bold: true }, { text: "把“评估”从拟合优度转向预测误差" }]));

// ============ Ch1 ============
content.push(H1("Ch1 · 线性回归模型基础"));
content.push(P("Essentials of Linear Regression Models", { run: { italics: true, color: "595959" } }));
content.push(Quote("📌 一句话：把 Normal / Logistic / Poisson 回归统一成 GLM 三件套（分布 + 线性预测 + link function），并复习 ML 估计与统计推断的整套语汇。"));

content.push(H2("1.1 课程概念地图 (p. 5)"));
content.push(bulletP([{ text: "全课程 5 大模块：数据生成机制 / GLM / 重采样与计算 / 正则化与灵活回归 / 特殊扩展" }]));

content.push(H2("1.2 为什么做回归 (Why regression) (p. 7-8)"));
content.push(bulletP([{ text: "回归在建模条件均值 " }, { text: "E[Y_i|X_i]", code: true }]));
content.push(MP([{ text: "🔴 三大用途：", bold: true, color: "C00000" }]));
content.push(bulletP([{ text: "理解数据生成过程（参数解释）" }]));
content.push(bulletP([{ text: "预测", bold: true }, { text: "（最小化 " }, { text: "E[l(Y, f̂(X))]", code: true }, { text: "）→ 链接 Ch5" }]));
content.push(bulletP([{ text: "因果推断", bold: true }, { text: " " }, { text: "E[Y|do X]", code: true }, { text: " → 链接 Ch4" }]));

content.push(H2("1.3 GLM 解剖 (Anatomy of a GLM) (p. 10-13)"));
content.push(MP([{ text: "🔴 GLM 三件套（必背）：", bold: true, color: "C00000" }]));
content.push(bulletP([{ text: "(a) 响应分布属于 " }, { text: "exponential family", bold: true }]));
content.push(bulletP([{ text: "(b) 解释变量的 " }, { text: "linear predictor", bold: true }, { text: " " }, { text: "Xβ", code: true }]));
content.push(bulletP([{ text: "(c) " }, { text: "Link function ", bold: true }, { text: "f(μ_i) = X_iβ", code: true }, { text: " 把均值连到线性组合" }]));
content.push(P(""));
content.push(MP([{ text: "四大假设：分布 / link 线性化 / 独立性 / 方差由分布决定" }]));

content.push(H2("1.4 统计建模与 ML 估计 (p. 15-18)"));
content.push(MP([{ text: "🔴 ML 估计三步：", bold: true, color: "C00000" }, { text: "写 log-likelihood → 求偏导 → 设零解 β" }]));
content.push(P("解析解 vs 数值解："));
content.push(bulletP([{ text: "Normal：", bold: true }, { text: "β̂ = (XᵀX)⁻¹XᵀY", code: true }, { text: "（解析）" }]));
content.push(bulletP([{ text: "Logistic / Poisson：", bold: true }, { text: "用 " }, { text: "IRLS", bold: true }, { text: " 数值迭代（iteratively re-weighted least squares）" }]));
content.push(MP([{ text: "📌 R 函数：", bold: true }, { text: "glm(formula, family=..., data=...)", code: true }, { text: " 底层就是 IRLS" }]));

content.push(H2("1.5 统计推断 (Statistical Inference) (p. 19-21)"));
content.push(bulletP([{ text: "推断三类对象：模型系数 / 给定 X 下 Y 的预测 / 哪些变量“需要”" }]));
content.push(bulletP([{ text: "Sampling distribution", bold: true }, { text: " 概念：想象重复抽样下 fitted model 的分布" }]));

content.push(H2("1.6 GLM 三大典例"));
content.push(makeTable(
  ["模型", "分布", "Link", "R 代码", "页码"],
  [
    [
      { text: "Normal (Catheter)", bold: true },
      "N(μ, σ²)",
      "identity",
      { text: 'glm(family="gaussian") 或 lm()', code: true },
      "24-30"
    ],
    [
      { text: "Logistic (CHD)", bold: true },
      "Binom(n,p)",
      { text: "logit log(p/(1-p))", bold: true },
      { text: 'glm(cbind(yes,no)~., family=binomial)', code: true },
      "32-46"
    ],
    [
      { text: "Poisson (Crab)", bold: true },
      "Poisson(μ)",
      { text: "log", bold: true },
      { text: 'glm(..., family=poisson)', code: true },
      "48-52"
    ]
  ],
  [2400, 1400, 1600, 3260, 700]
));
content.push(P(""));
content.push(MP([{ text: "🔴 Logistic CI 两种做法：", bold: true, color: "C00000" }, { text: "直接做（粗糙）vs 在 linear predictor 上做 CI 再变换（更稳，p. 41-42）" }]));
content.push(MP([{ text: "🔴 Probit link：", bold: true, color: "C00000" }, { text: "p = Φ(Xβ)，与 logit 在中心区差异小（p. 43-46）" }]));
content.push(MP([{ text: "📌 R 函数：", bold: true }, { text: 'predict(fit, newdata, type="response"/"link", se.fit=TRUE)', code: true }, { text: "、" }, { text: 'factor() + contrasts()', code: true }]));

content.push(H2("1.7 Appendix (p. 54-59)"));
content.push(bulletP([{ text: "log / logit / expit / odds 代数恒等式（", bold: false }, { text: "重点公式", bold: true, color: "C00000" }, { text: "）" }]));
content.push(bulletP([{ text: "矩阵记号统一：" }, { text: "Y, μ, Σ_Y, Xβ", code: true }]));

// ============ Ch2 ============
content.push(H1("Ch2 · 模型诊断与拟合优度"));
content.push(P("Diagnostics & Goodness of Fit", { run: { italics: true, color: "595959" } }));
content.push(Quote("📌 一句话：模型拟出来怎么判好坏？两条线 ——诊断找出问题点（point-level），GoF / Deviance 给出模型整体评分（model-level），最后用 dispersion 修正分布假设错配。"));

content.push(H2("2.1 诊断的两种问题 (p. 3-5)"));
content.push(bulletP([{ text: "数据问题", bold: true }, { text: "：outliers / 高杠杆点 / 多重共线" }]));
content.push(bulletP([{ text: "模型问题", bold: true }, { text: "：linearity / Normality / 独立同方差" }]));

content.push(H2("2.2 Residuals (p. 7-9)"));
content.push(MP([{ text: "🔴 Pearson 残差：", bold: true, color: "C00000" }, { text: " r_i = (Y_i - μ̂_i) / √Var(Y_i|X_i)", code: true }]));
content.push(bulletP([{ text: "r_i > 0", code: true }, { text: "：模型低估；" }, { text: "r_i < 0", code: true }, { text: "：模型高估" }]));
content.push(bulletP([{ text: "Binomial 形式：" }, { text: "r_i = (Y_i - n_i·p̂_i) / √(n_i·p̂_i·(1-p̂_i))", code: true }]));

content.push(H2("2.3 影响点 (Influential points) (p. 11-14)"));
content.push(bulletP([{ text: "Cook's distance", bold: true }, { text: "：移除第 i 点后 β̂ 的整体变动" }]));
content.push(bulletP([{ text: "Leverage h_ii", bold: true }, { text: "：hat matrix " }, { text: "H = X(XᵀX)⁻¹Xᵀ", code: true }, { text: " 的对角元" }]));
content.push(MP([{ text: "📌 R 函数：", bold: true }, { text: "plot(glm_fit)", code: true }, { text: " 自动出 4 张诊断图" }]));

content.push(H2("2.4 拟合优度: Deviance (p. 17-18)"));
content.push(MP([{ text: "🔴 核心定义：", bold: true, color: "C00000" }, { text: "Dev = 2·logL_sat − 2·logL_mod", code: true }]));
content.push(bulletP([{ text: "Saturated model", bold: true }, { text: "：每个 covariate pattern 单独拟到 μ̂_i = y_i" }]));
content.push(bulletP([{ text: "小 deviance = 接近最优拟合" }]));

content.push(H2("2.5 F-test —— Normal 回归专属 (p. 20-26)"));
content.push(MP([{ text: "🔴 Added variable F-test：", bold: true, color: "C00000" }]));
content.push(...codeBlock([
  "f_o = [(RSS_S − RSS_F)/(k_F − k_S)] / [RSS_F/(n − k_F − 1)]"
]));
content.push(MP([{ text: "📌 R 函数：", bold: true }, { text: "anova(fit1, fit2)", code: true }, { text: "、" }, { text: "summary(lm_fit)", code: true }, { text: " 底部那行 F-statistic" }]));

content.push(H2("2.6 χ² 检验 —— Logistic / Poisson (p. 28-38)"));
content.push(bulletP([{ text: "Pearson χ² 拟合优度统计量；近似有效条件：每格 " }, { text: "r_i ≥ 5", code: true }, { text: " 且 " }, { text: "n_i - r_i ≥ 5", code: true }]));
content.push(bulletP([{ text: "Grouped vs ungrouped data 概念（p. 34-35）" }]));

content.push(H2("2.7 Likelihood Ratio Test (p. 40-47)"));
content.push(MP([{ text: "🔴 LRT 统计量：", bold: true, color: "C00000" }, { text: " d_o = Dev_sub − Dev_full ~ χ²_v", code: true }, { text: "，v = 多出的变量数" }]));
content.push(Warn("⚠️ ANOVA 顺序敏感：相同变量按不同顺序加入，p 值会不同（见 crab 例 p. 43-46）"));
content.push(MP([{ text: "📌 R 函数：", bold: true }, { text: 'anova(glm_fit, test="Chisq")', code: true }]));

content.push(H2("2.8 过离散与 Dispersion 建模 🔴 重点 (p. 49-64)"));
content.push(makeTable(
  ["方案", "形式", "优点", "缺点", "R 函数"],
  [
    [
      { text: "Quasi-likelihood", bold: true },
      "Var(Y) = τμ（Poisson）",
      "过/欠离散都行",
      "不能做 deviance/GoF",
      { text: "family=quasipoisson / quasibinomial", code: true }
    ],
    [
      { text: "Negative Binomial", bold: true },
      "Var = μ + μ²/θ",
      "保留 likelihood",
      "仅处理过离散",
      { text: "MASS::glm.nb()", code: true }
    ]
  ],
  [1800, 1800, 1600, 1700, 2460]
));
content.push(P(""));
content.push(bulletP([{ text: "Quasi 模型 anova 推荐用 " }, { text: 'test="F"', code: true }, { text: " 而非 " }, { text: 'Chisq', code: true }, { text: "（p. 54）" }]));
content.push(bulletP([{ text: "过离散两大原因：缺失重要变量 / 观测不独立" }]));

content.push(H2("2.9 Offsets (p. 66-72)"));
content.push(bulletP([{ text: "处理速率数据（如死亡率 / 10 万人口）" }]));
content.push(bulletP([{ text: "把 " }, { text: "log(x/100000)", code: true }, { text: " 作为系数固定为 1 的 regressor" }]));
content.push(MP([{ text: "📌 R 代码：", bold: true }, { text: "glm(n ~ x + offset(log(pop/100000)), family=poisson)", code: true }]));

// ============ Ch3 ============
content.push(H1("Ch3 · 模型选择"));
content.push(P("Model Selection", { run: { italics: true, color: "595959" } }));
content.push(Quote("📌 一句话：Ch2 的 F-test / LRT 一次只能比两个嵌套模型，本章引入信息准则 (AIC/AICc/BIC) 比较任意模型并搜索最优组合。"));

content.push(H2("3.1 模型选择问题 (p. 3-4)"));
content.push(bulletP([{ text: "p 个变量 → 2^p 候选模型；deviance test 局限：只能两两 + 必须嵌套" }]));

content.push(H2("3.2 信息论方法优势 (p. 5)"));
content.push(bulletP([{ text: "可比较" }, { text: "非嵌套模型", bold: true }, { text: "，可一次评估所有候选" }]));

content.push(H2("3.3 三大信息准则 🔴 必背公式 (p. 17-25)"));
content.push(makeTable(
  ["准则", "公式", "偏好", "何时用"],
  [
    [
      { text: "AIC", bold: true },
      { text: "-2ℓ + 2k", code: true },
      "较复杂",
      "大样本 + 关心预测损失"
    ],
    [
      { text: "AICc", bold: true },
      { text: "-2ℓ + 2k + (2k²+2k)/(n-k-1)", code: true },
      "同 AIC",
      "小样本（首选）"
    ],
    [
      { text: "BIC", bold: true },
      { text: "-2ℓ + k·log(n)", code: true },
      "较简单",
      "想找“真模型”"
    ]
  ],
  [1200, 3500, 1500, 3160]
));
content.push(P(""));
content.push(Warn("⚠️ 警告（p. 26）：信息准则是相对度量；响应不同（含变换）的模型不能比"));

content.push(H2("3.4 GAM 用作非线性诊断 (p. 11-13)"));
content.push(bulletP([{ text: "配主模型外用 GAM 看 partial effect 是否非线性" }]));
content.push(bulletP([{ text: "不能用直线穿过 dotted band → 考虑加 quadratic" }]));
content.push(MP([{ text: "📌 R 函数：", bold: true }, { text: "mgcv::gam()", code: true }]));

content.push(H2("3.5 模型搜索策略 (p. 28-37)"));
content.push(makeTable(
  ["策略", "R 函数", "适用"],
  [
    [
      { text: "Exhaustive", bold: true },
      { text: "MuMIn::dredge(full_model)", code: true },
      "优先选；变量数小时"
    ],
    [
      { text: "Backward", bold: true },
      { text: 'step(full, direction="backward")', code: true },
      "大模型空间"
    ],
    [
      { text: "Forward", bold: true },
      { text: 'step(null, direction="forward")', code: true },
      "想从空模型探索"
    ],
    [
      { text: "Both", bold: true },
      { text: 'step(full, direction="both")', code: true },
      "兼顾加减"
    ]
  ],
  [1600, 4200, 3560]
));
content.push(P(""));
content.push(MP([{ text: "🔴 必备前置：", bold: true, color: "C00000" }, { text: 'options(na.action = "na.fail")', code: true }]));
content.push(MP([{ text: "📌 辅助：", bold: true }, { text: "get.models(dredge_obj, 1)[[1]]", code: true }, { text: " 取最佳模型，" }, { text: "plot(dredge_obj)", code: true }, { text: " 看变量入选频率" }]));

content.push(H2("3.6 警示 (p. 38)"));
content.push(bulletP([{ text: "信息准则不替代领域知识" }]));

// ============ Ch4 ============
content.push(H1("Ch4 · 因果推断"));
content.push(P("Causal Models", { run: { italics: true, color: "595959" } }));
content.push(Quote("📌 一句话：前三章在做“X 与 Y 关联吗”，本章升级到“X 真的导致 Y 吗” —— 用 counterfactual 定义 causation，用 causal diagram (DAG) 画因果假设，再选合适的 regressors 估计直接 / 总因果效应。"));

content.push(H2("4.1 因果建模的两条路 (p. 3-6)"));
content.push(bulletP([{ text: "Designed experiment (RCT)", bold: true }, { text: "：gold standard，随机分配处理" }]));
content.push(bulletP([{ text: "Causal analysis of observational data", bold: true }, { text: "：本章重点" }]));
content.push(bulletP([{ text: "Counterfactual", bold: true }, { text: " 概念：同一观测在 do(X=1) vs do(X=0) 的输出差" }]));

content.push(H2("4.2 因果性的数学定义 (p. 8-11)"));
content.push(bulletP([{ text: "Individual causal effect: " }, { text: "Y_i^{a=1} - Y_i^{a=0}", code: true }]));
content.push(MP([{ text: "🔴 Average Causal Effect (ACE): ", bold: true, color: "C00000" }, { text: "ACE = E(Y^{a=1}) - E(Y^{a=0})", code: true }]));
content.push(MP([{ text: "根本难题：", bold: true }, { text: "一个人只能观察到一种 outcome" }]));

content.push(H2("4.3 观察数据的麻烦 (p. 13-14)"));
content.push(bulletP([{ text: "self-selection → group 间初始差异 → 简单组间差 ≠ ACE" }]));

content.push(H2("4.4 Causal Diagram (DAG) (p. 16-25)"));
content.push(bulletP([{ text: "节点 = 变量，箭头 = 因果关系" }]));
content.push(bulletP([{ text: "相关关系 5 大来源：抽样误差 / A→B / B→A / 共因 C / 选择偏差" }]));
content.push(bulletP([{ text: "历史案例：Fisher vs BMJ 关于吸烟致癌的争论（p. 19-23）" }]));
content.push(P(""));
content.push(MP([{ text: "建立因果的 3 原则（p. 25）：", bold: true }]));
content.push(bulletP([{ text: "因 在 果 之前" }]));
content.push(bulletP([{ text: "必须有关联（且关联越强越好）" }]));
content.push(bulletP([{ text: "关联不能被第三变量解释" }]));

content.push(H2("4.5 路径四分类 🔴 全章核心，必背 (p. 32-37, 70-72)"));
content.push(makeTable(
  ["路径", "图示", "开/关", "估直接效应", "估总效应"],
  [
    [
      { text: "Direct（直接因）", bold: true },
      "S → R",
      "开",
      "保留",
      "保留"
    ],
    [
      { text: "Indirect（间接因）", bold: true },
      "S → A → R",
      "开",
      { text: "关掉 A", bold: true, color: "C00000" },
      "保留"
    ],
    [
      { text: "Confounding（混杂）", bold: true },
      "S ← A → R",
      "开",
      { text: "关掉 A", bold: true, color: "C00000" },
      { text: "关掉 A", bold: true, color: "C00000" }
    ],
    [
      { text: "Colliding（碰撞）", bold: true },
      "S → A ← R",
      "天然关",
      { text: "不含 A！", bold: true, color: "C00000" },
      { text: "不含 A！", bold: true, color: "C00000" }
    ]
  ],
  [2200, 1800, 1100, 2130, 2130]
));
content.push(P(""));
content.push(MP([{ text: "🔴 操作要点：", bold: true, color: "C00000" }]));
content.push(bulletP([{ text: "“关闭路径” = 把路径上某变量加进模型当 regressor" }]));
content.push(bulletP([{ text: "混杂变量", bold: true }, { text: "必加", bold: true, color: "C00000" }, { text: "；Collider " }, { text: "绝不能加", bold: true, color: "C00000" }]));
content.push(bulletP([{ text: "Effect modifier", bold: true }, { text: "（交互项）在 DAG 里看不出来，要用 interaction term 测" }]));

content.push(H2("4.6 操作准则 (p. 40-44)"));
content.push(MP([{ text: "估计直接效应", bold: true }, { text: " → 把所有“直接影响 response 的变量”全放进模型即可" }]));
content.push(MP([{ text: "估计总效应", bold: true }, { text: " → 保留间接路径（不加中介），只关混杂" }]));

content.push(H2("4.7 经典案例"));
content.push(bulletP([{ text: "College plans 例", bold: true }, { text: "（p. 47-69）：SES, IQ, PE 对 CP 的直接 vs 总效应" }]));
content.push(bulletP([{ text: "Salary gap 例", bold: true }, { text: "（p. 73-92）：UW 工资性别诉讼；性别的直接效应 (~$300/月) vs 总效应 (~$1100/月) 差距巨大" }]));

content.push(H2("4.8 漏变量的后果 (p. 99-107)"));
content.push(makeTable(
  ["漏掉的类型", "直接效应估计", "总效应估计"],
  [
    ["漏混杂", { text: "有偏", color: "C00000", bold: true }, { text: "有偏", color: "C00000", bold: true }],
    ["漏 collider", { text: "不漏 OK（本来就不该加）", color: "2E7D32" }, { text: "同上", color: "2E7D32" }],
    ["漏间接路径中介", { text: "有偏", color: "C00000", bold: true }, { text: "无偏", color: "2E7D32", bold: true }],
    ["漏单纯直接效应变量", "无偏但精度低", "同上"]
  ],
  [3000, 3180, 3180]
));

// ============ Ch5 ============
content.push(H1("Ch5 · 预测、交叉验证与自助法"));
content.push(P("Prediction, CV, Bootstrap", { run: { italics: true, color: "595959" } }));
content.push(Quote("📌 一句话：Ch1-Ch4 关注拟合优度（用同一份数据训练并评判），本章把目光转向预测 —— 用 K-fold CV 客观估计模型在新数据上的表现，用 Bootstrap 重采样得到任意统计量的经验抽样分布。"));

content.push(Warn("⚠️ 期中范围提示：老师课上明确说，期中只考到本章的 cross-validation 部分，bootstrap 不在期中范围。期末则全章都考。"));

content.push(H2("5.1 预测的根本问题 (p. 3-5)"));
content.push(bulletP([{ text: "“Prediction is difficult, especially when dealing with the future”", italics: true }]));
content.push(bulletP([{ text: "Model assessment", bold: true }, { text: "：估计模型在新数据上的预测误差" }]));
content.push(bulletP([{ text: "Model selection", bold: true }, { text: "：基于预测表现选最优模型" }]));

content.push(H2("5.2 Loss Function (p. 7-9)"));
content.push(bulletP([{ text: "风险定义：" }, { text: "R(θ) = E_{X,Y}[L(Y, f_θ(X))]", code: true }, { text: "，经验估计 " }, { text: "R̂ = (1/n)·Σ L(y_i, f_θ(x_i))", code: true }]));
content.push(bulletP([{ text: "连续 Y", bold: true }, { text: ": squared / absolute / log-likelihood loss" }]));
content.push(bulletP([{ text: "分类 Y", bold: true }, { text: ": 0-1 loss / log-likelihood loss" }]));

content.push(H2("5.3 K-fold Cross-Validation (p. 17-21)"));
content.push(MP([{ text: "🔴 算法：", bold: true, color: "C00000" }]));
content.push(bulletP([{ text: "n 个观测均分为 K 折 S_1, ..., S_K" }]));
content.push(bulletP([{ text: "每次留 1 折做 test，其余 K-1 折 train" }]));
content.push(bulletP([{ text: "平均：" }, { text: "R̂_K = (1/K)·Σ e_k", code: true }]));
content.push(P("K 的选择：太小 bias 高，太大 variance 高（bias-variance trade-off）"));

content.push(H2("5.4 Leave-One-Out CV (LOOCV) (p. 23-25)"));
content.push(bulletP([{ text: "极端情形 K = n，每次留一个点测试" }]));
content.push(bulletP([{ text: "bias 几乎 0 但 variance 高（折间相关 ρ ≈ 1）" }]));

content.push(H2("5.5 CV 的理论性质 (p. 28-33)"));
content.push(bulletP([{ text: "一致性", bold: true }, { text: "：固定 K, n → ∞ 时 " }, { text: "R̂_K →ᵖ 真风险", code: true }]));
content.push(bulletP([{ text: "Bias", bold: true }, { text: " = O(1/(nK))；K 大 → bias 小" }]));
content.push(bulletP([{ text: "Variance", bold: true }, { text: " ≈ σ²·(1/K + (K-1)·ρ/K)；K 大 → variance 大（因为 ρ 大）" }]));
content.push(Warn("⚠️ Var 没有通用无偏估计（Bengio & Grandvalet 2004）"));

content.push(P(""));
content.push(MP([{ text: "🔴 Bias-Variance Trade-off（必背）：", bold: true, color: "C00000" }]));
content.push(makeTable(
  ["K", "Bias", "Variance"],
  [
    ["Small K", { text: "High", color: "C00000" }, { text: "Low", color: "2E7D32" }],
    ["Large K (LOOCV)", { text: "Low", color: "2E7D32" }, { text: "High", color: "C00000" }]
  ],
  [3120, 3120, 3120]
));

content.push(H2("5.6 Bootstrap (Non-Parametric) (p. 35-46)"));
content.push(bulletP([{ text: "思想：用" }, { text: "有放回重采样", bold: true }, { text: "代替重新抽样" }]));
content.push(MP([{ text: "🔴 算法：", bold: true, color: "C00000" }]));
content.push(bulletP([{ text: "从 Z_1, ..., Z_n 有放回采样得 Z*ᵇ（大小同 n）" }]));
content.push(bulletP([{ text: "在 Z*ᵇ 算 S(Z*ᵇ)" }]));
content.push(bulletP([{ text: "重复 B 次（通常 B ≥ 1000）" }]));
content.push(bulletP([{ text: "用 B 个 S 的经验分布做推断" }]));
content.push(P(""));
content.push(MP([{ text: "常用度量：", bold: true }]));
content.push(bulletP([{ text: "Bias = " }, { text: "S̄* - Ŝ_n", code: true }]));
content.push(bulletP([{ text: "Variance = " }, { text: "(1/B)·Σ(S(Z*ᵇ) - S̄*)²", code: true }]));
content.push(bulletP([{ text: "CI = 排序后取 α/2 分位数" }]));
content.push(P(""));
content.push(MP([{ text: "何时用（p. 45-46）：", bold: true }, { text: "理论难推；理论假设可疑；想避开分布误设" }]));
content.push(Warn("⚠️ 注意：bootstrap 不是为了得到更好的点估计；要求数据 iid；不能修正模型 misspecification"));

// ============ End: review path ============
content.push(H1("复习路线建议"));

content.push(bulletP([{ text: "第一遍：", bold: true }, { text: "先把 Ch1 GLM 三件套和 Ch4 路径四分类背到能默写" }]));
content.push(bulletP([{ text: "第二遍：", bold: true }, { text: "在 Ch2-3 串起 deviance → AIC/BIC 的逻辑链，理解为什么后者比前者强" }]));
content.push(bulletP([{ text: "第三遍：", bold: true }, { text: "Ch5 的 CV 与前面 AIC 形成对照（理论 vs 经验），bootstrap 单独看" }]));
content.push(bulletP([{ text: "R 代码：", bold: true }, { text: "建议自己跑一遍 catheter / CHD / crab / wine quality / salary 五个案例" }]));

content.push(P(""));
content.push(P("Ch6-Ch10 留到下个会话再处理。", { run: { italics: true, color: "595959" } }));

// ============ Document ============
const doc = new Document({
  creator: "STATS 762 Revision",
  title: "STATS 762 期末复习大纲 (Ch1-Ch5)",
  styles: {
    default: {
      document: { run: { font: FONT, size: 22 } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: FONT, color: "1F4E79" },
        paragraph: { spacing: { before: 240, after: 180 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: FONT, color: "2E74B5" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: FONT, color: "1F4E79" },
        paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 }
      }
    ]
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          },
          {
            level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1440, hanging: 360 } } }
          }
        ]
      }
    ]
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: "STATS 762 · 期末复习大纲 (Ch1-Ch5)", font: FONT, size: 18, color: "808080" })]
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "第 ", font: FONT, size: 18, color: "808080" }),
              new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: "808080" }),
              new TextRun({ text: " 页 / 共 ", font: FONT, size: 18, color: "808080" }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: "808080" }),
              new TextRun({ text: " 页", font: FONT, size: 18, color: "808080" })
            ]
          })]
        })
      },
      children: content
    }
  ]
});

const outputPath = "C:\\Users\\33238\\Desktop\\2026 semester1 选课\\2.STATS 762\\STATS762_期末复习大纲_Ch1-Ch5.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log("✓ Generated:", outputPath);
  console.log("  Size:", (buffer.length / 1024).toFixed(1), "KB");
});
