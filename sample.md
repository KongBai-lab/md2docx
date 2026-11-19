# 公式排版示例

[[TOC]]

本文档演示 **文字与公式分离**：正文是普通文字，公式是 Office 原生公式对象（OMML），在 WPS / Word 里可以双击编辑，也可以右键「转换为 LaTeX 代码」查看源码。

## 1. 行内公式

质能方程 $E = mc^2$ 说明质量与能量等价；二次方程 $x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$ 的判别式是 $\Delta = b^2 - 4ac$。

## 2. 多行公式

麦克斯韦方程组（多行对齐）：

$$
\begin{aligned}
\nabla \cdot \boldsymbol{E} &= \frac{\rho}{\varepsilon_0} \\
\nabla \cdot \boldsymbol{B} &= 0 \\
\nabla \times \boldsymbol{E} &= -\frac{\partial \boldsymbol{B}}{\partial t} \\
\nabla \times \boldsymbol{B} &= \mu_0 \left( \boldsymbol{J} + \varepsilon_0 \frac{\partial \boldsymbol{E}}{\partial t} \right)
\end{aligned}
$$

## 3. 分段函数与矩阵

$$
f(x) = \begin{cases} x^2 & x \ge 0 \\ -x & x < 0 \end{cases}
$$

$$
A = \begin{pmatrix} a & b \\ c & d \end{pmatrix}, \quad \det A = ad - bc
$$

## 4. 列表与表格

- 求和公式（独立成行显示，上下限位置才正确）：

$$
\sum_{i=1}^{n} i = \frac{n(n+1)}{2}
$$

- 极限公式（行内）：$\lim_{x \to 0} \frac{\sin x}{x} = 1$

| 符号 | 含义 |
| --- | --- |
| $\alpha$ | 角度 |
| $\nabla$ | 梯度算子 |

## 5. 正文中的符号与特殊字母

说明文字里同样可以直接写符号，不需要每次都套公式：设 α 为入射角、θ 为折射角，∂u/∂t 表示时间偏导数，∇ 表示梯度算子，∑ 与 ∏ 分别表示求和与连乘，∫ 表示积分。

常见符号一览：Δ Ω ∂ ∇ ∑ ∏ ∫ ∮ ∞ ≈ ≠ ≤ ≥ ⩽ ⩾ ± × ÷ √ → ⇒ ⇔ ∈ ∉ ⊂ ⊆ ∪ ∩ ∅ ∀ ∃ ¬ ∧ ∨ ∴ ∵ ° ′ ″ ⟨ ⟩ ⌈ ⌉ ⌊ ⌋ ℓ ℏ ℝ ℕ ℤ ℚ ℂ。

带向量与均值标记的拉丁字母也可以直接写在正文里：速度 v⃗、加速度 a⃗、力 F⃗ 是向量；样本均值 x̄、ȳ 与总体均值 μ 使用横线标记。组合示例：x̄ = (x₁ + x₂ + … + xₙ)/n。

也可以直接写 LaTeX 命令（不用 `$`）：\vec{v}、\bar{x}、\hat{x}，生成时会转成同样的行内公式对象，和公式区的字形完全一致。

连续书写的整段表达式会合并成一个公式对象：E = mc²、F = ma、a + b = c。

$$
\vec{v} = \frac{d\vec{r}}{dt}, \qquad \bar{x} = \frac{1}{n}\sum_{i=1}^{n} x_i
$$

| 符号 | 名称 | 符号 | 名称 |
| --- | --- | --- | --- |
| α | 阿尔法 | β | 贝塔 |
| θ | 西塔 | λ | 兰姆达 |
| ∇ | 梯度 | ∂ | 偏导 |
| ∑ | 求和 | ∫ | 积分 |
| v⃗ | 速度向量 | F⃗ | 力向量 |
| x̄ | 样本均值 | ȳ | 均值 |

## 6. 引用与代码

> 公式对象在 WPS 中可用「公式 → 转换 → 转换为 LaTeX 代码」查看源码。

```python
print("hello")
```

## 7. 麦克斯韦方程组（积分形式）

$$
\begin{aligned}
\oint_{\partial V} \boldsymbol{E} \cdot d\boldsymbol{A} &= \frac{Q_{\text{enc}}}{\varepsilon_0} \\
\oint_{\partial V} \boldsymbol{B} \cdot d\boldsymbol{A} &= 0 \\
\oint_{\partial S} \boldsymbol{E} \cdot d\boldsymbol{l} &= -\frac{d\Phi_{\boldsymbol{B}}}{dt} \\
\oint_{\partial S} \boldsymbol{B} \cdot d\boldsymbol{l} &= \mu_0 I_{\text{enc}} + \mu_0 \varepsilon_0 \frac{d\Phi_{\boldsymbol{E}}}{dt}
\end{aligned}
$$

## 8. 定积分与多重积分

$$
\int_{0}^{1} x^2\,dx = \frac{1}{3}, \qquad \int_{-\infty}^{\infty} e^{-x^2}\,dx = \sqrt{\pi}
$$

$$
\iint_{D} (x^2 + y^2)\,dx\,dy, \qquad \iiint_{V} \rho\,dV, \qquad \oint_{C} \boldsymbol{F} \cdot d\boldsymbol{l}
$$

$$
\iint_{S} \boldsymbol{F} \cdot d\boldsymbol{S} = \iiint_{V} \nabla \cdot \boldsymbol{F}\,dV
$$

## 9. 极限、求和与连乘

$$
\lim_{x \to 0} \frac{\sin x}{x} = 1, \qquad
\sum_{i=1}^{n} i = \frac{n(n+1)}{2}, \qquad
\prod_{k=1}^{n} k = n!
$$

$$
\sum_{i=1}^{n} i^2 = \frac{n(n+1)(2n+1)}{6}, \qquad
\int_{-\infty}^{\infty} e^{-x^2}\,dx = \sqrt{\pi}
$$

## 10. 图片与超链接

![示例图片](out/test.png)

图 1 示例图片（独立成段时居中）。

下图是混排示例：![示例图片](out/test.png) 图 2 混排图片（题注居中，五号字）。

![示例图片](out/test.png)
图 3 图片与题注同段时也会自动拆分。

示例站点是 [示例站点](https://example.com)。



