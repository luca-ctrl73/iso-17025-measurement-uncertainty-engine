# UncertLab 17025 Pro: Deterministic Measurement Uncertainty Engine for ISO/IEC 17025:2017 Compliance

[![ISO/IEC 17025:2017](https://img.shields.io/badge/Compliance-ISO%2FIEC%2017025%3A2017-0052CC.svg)](#)
[![JCGM 100:2008 GUM](https://img.shields.io/badge/Standard-JCGM%20100%3A2008%20(GUM)-008080.svg)](#)
[![ILAC-G8:09/2019](https://img.shields.io/badge/Decision%20Rules-ILAC--G8%3A09%2F2019-2E8B57.svg)](#)
[![Architecture](https://img.shields.io/badge/Engine-Macro--Free%20OpenXML%20(.xlsx)-217346.svg)](#)
[![Security](https://img.shields.io/badge/InfoSec-100%25%20Offline%20Data%20Sovereignty-darkgreen.svg)](#)

**UncertLab 17025 Pro** is an audit-ready, deterministic measurement uncertainty calculation system engineered by **BioTechExplorers** for testing and calibration laboratories accredited (or pursuing accreditation) under **ISO/IEC 17025:2017**.

Operating exclusively on native OpenXML formulas within Microsoft Excel (`.xlsx`), UncertLab completely eliminates VBA macros, third-party binary dependencies, and external cloud connections, ensuring total cybersecurity compliance, offline data sovereignty, and unconstrained audit defensibility during ACCREDIA, UKAS, DAkkS, A2LA, and COFRAC assessments.

---

## 1. Quick Access & Deployment Links

* **Free Bench Reference Card (0 € Lead Magnet):**  
  [Download the ISO/IEC 17025 Metrology Bench QuickRef Card](https://biotechexplorers.gumroad.com/l/17025-bench-card)
* **Master Calculation Engine & Audit Validation Dossiers:**  
  [Access the Official UncertLab 17025 Pro Suite](https://biotechexplorers.gumroad.com/l/uncertlab-17025-pro)

---

## 2. Metrological Architecture & Standard Compliance

UncertLab 17025 Pro is engineered strictly in compliance with international metrological standards:

| Standard / Guide | Scope & Implementation in UncertLab |
| :--- | :--- |
| **ISO/IEC 17025:2017 (§7.6)** | Systematic evaluation of measurement uncertainty for all accredited testing and calibration methods. |
| **JCGM 100:2008 (GUM)** | Evaluation of measurement data — Guide to the expression of uncertainty in measurement. |
| **ILAC-G8:09/2019** | Decision rules and conformance reporting with binary acceptance intervals using guard bands ($w = U$). |
| **Welch-Satterthwaite** | Effective degrees of freedom ($\nu_{\text{eff}}$) calculation for dynamic coverage factor ($k$) determination. |
| **ISO 5725-2** | Accuracy (trueness and precision) of measurement methods and results. |
| **ASTM E178 / Grubbs** | Bilateral outlier detection shields preventing division-by-zero (`#DIV/0!`) system halts. |

---

## 3. Mathematical Formulations Implemented

### A. Combined Standard Uncertainty (Correlated 5x5 Covariance Matrix)
Under JCGM 100:2008 clause 5.2.2, the combined standard uncertainty $u_c(y)$ for input quantities $x_i$ with known correlation coefficients $r(x_i, x_j)$ is computed via:

$$u_c^2(y) = \sum_{i=1}^{N} c_i^2 u^2(x_i) + 2 \sum_{i=1}^{N-1} \sum_{j=i+1}^{N} c_i c_j u(x_i) u(x_j) r(x_i, x_j)$$

Where $c_i = \frac{\partial f}{\partial x_i}$ represents the partial derivative sensitivity coefficient.

### B. Welch-Satterthwaite Effective Degrees of Freedom
To prevent the non-conservative assumption of $k = 2.00$ when Type A repeatability components have limited sample size ($n < 30$), the effective degrees of freedom $\nu_{\text{eff}}$ are derived deterministically:

$$\nu_{\text{eff}} = \frac{u_c^4(y)}{\sum_{i=1}^{N} \frac{c_i^4 u^4(x_i)}{\nu_i}}$$

The coverage factor $k$ is dynamically extracted using Student's t-distribution for a 95.45% bilateral confidence interval:

$$k = t_{0.9545}(\nu_{\text{eff}})$$

### C. Type B Analytical Probability Distribution Divisors
* **Rectangular (Uniform):** $u(x_i) = \frac{a}{\sqrt{3}} \approx \frac{a}{1.73205}$ (instrument tolerances, ambient drift)
* **Triangular (Symmetric):** $u(x_i) = \frac{a}{\sqrt{6}} \approx \frac{a}{2.44949}$ (volumetric dispensing, temperature fluctuation bounds)
* **Digital Resolution Step:** $u(x_i) = \frac{d}{2\sqrt{3}} \approx \frac{d}{3.46410}$ (least significant digit display resolution)
* **Normal / Gaussian:** $u(x_i) = \frac{U}{k}$ (calibration certificates at $k=1$ or $k=2$)

### D. ILAC-G8:09/2019 Guard Band Decision Logic
Binary acceptance intervals configured to mitigate shared risk:
* **Conforming (Pass):** $\text{Measurement Result} + U \le \text{Upper Specification Limit (USL)}$
* **Non-Conforming (Fail):** $\text{Measurement Result} - U > \text{Lower Specification Limit (LSL)}$
* **Conditional / Guard Band Region:** $\vert{}\text{Result} - \text{Limit}\vert{} < U$ (requires documented client contractual agreement under §7.8.6)

---

## 4. Engineering Comparison: UncertLab 17025 Pro vs. Alternative Solutions

| Capability / Parameter | In-House Ad-Hoc Sheets | Enterprise Metrology Software (e.g., GUM Workbench) | UncertLab 17025 Pro |
| :--- | :--- | :--- | :--- |
| **Cost Model** | Internal labor (>€1,000) | Recurring SaaS (€800–€2,500/year) | **One-time commercial license (€49–€189)** |
| **Macros & Code Security** | Often uses unvetted VBA macros | Proprietary executable binary (.exe) | **100% Native OpenXML (`.xlsx`), Zero Macros** |
| **IT InfoSec Friction** | Antivirus / quarantine risks | Requires IT enterprise installation permits | **Zero installation required (Runs natively on Excel)** |
| **Audit Qualification Logs** | Rarely documented | Vendor-managed | **Pre-compiled IQ/OQ/PQ Dossiers (.docx + .pdf)** |
| **Data Sovereignty** | Local | Cloud/Desktop vendor locked | **100% Offline, internal server isolation** |

---

## 5. Software Qualification & Validation (IQ / OQ / PQ)

Under ISO/IEC 17025:2017 clause 7.7 and clause 7.11 (*Control of data and information management*), spreadsheet calculations must be validated prior to routine testing.

UncertLab 17025 Pro includes a comprehensive **8-Page A4 Software Delivery & Validation Dossier**:
* **Installation Qualification (IQ):** Checksum integrity verification, hardware/software operating boundaries (Excel 2016 through Microsoft 365, LibreOffice Calc).
* **Operational Qualification (OQ):** Stress testing with zero variance inputs, Grubbs division-by-zero mitigation, and Welch-Satterthwaite edge cases.
* **Performance Qualification (PQ):** Numerical benchmark verification comparing UncertLab engine results against verified JCGM 100 GUM reference examples.
* **Auditor Sign-off Register:** Formal sign-off sections ready for technical assessors and Quality Managers.

---

## 6. Official Resources & Contact

* **Repository Maintainer:** BioTechExplorers Metrology Engineering Team
* **Technical Inquiries:** `biotech.explorer.work@gmail.com`
* **Direct Product Portal:** [https://biotechexplorers.gumroad.com/l/uncertlab-17025-pro](https://biotechexplorers.gumroad.com/l/uncertlab-17025-pro)
* **Free Resource Download:** [https://biotechexplorers.gumroad.com/l/17025-bench-card](https://biotechexplorers.gumroad.com/l/17025-bench-card)
