---
name: using-agent-skills
description: Discovers and invokes agent skills. Use when starting a session or when you need to discover which skill applies to the current task. This is the meta-skill that governs how all other skills are discovered and invoked.
---

# Using Agent Skills (Solution Architecture & Business Analysis Skill Registry)

## Overview

Agent Skills là kho vũ khí chuẩn mực enterprise gồm **53 Singleton Skills** tọa lạc tại thư mục `.agents/skills/`, được tổ chức chặt chẽ theo các phân tầng phát triển phần mềm Enterprise: từ **Business Analysis & Inception**, **Architecture & Design**, **High-Quality Implementation**, cho đến **Diagnostics, Verification & Operations**.

Mỗi skill tuân thủ nghiêm ngặt **Nguyên lý Trách nhiệm Đơn nhất (Singleton Responsibility Principle)**: mỗi file đảm nhiệm một vai trò duy nhất, không pha trộn các miền tri thức khác nhau.

---

## ⚡ NGUYÊN TẮC NGƯỠNG TỰ TIN 100% & QUY TRÌNH 4 BƯỚC BẮT BUỘC

> [!IMPORTANT]
> **QUY TRÌNH 4 BƯỚC BẤT DI BẤT DỊCH KHI NHẬN YÊU CẦU TỪ USER:**
> 
> 1. **Bước 1: Phỏng vấn 1-1 qua [`interview-me`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/interview-me/SKILL.md) đến khi đạt 100% Tự Tin**:
>    - Đối với mọi yêu cầu/prompt từ người dùng, **TUYỆT ĐỐI CẤM TỰ Ý ĐOÁN MÒ HOẶC VIẾT CODE NGAY**.
>    - Phỏng vấn tương tác từng câu hỏi một (one-question-at-a-time) kèm `HYPOTHESIS`, `CONFIDENCE: X%`, `Q:` và `GUESS:`.
>    - Phỏng vấn cho đến khi đạt mức **CONFIDENCE: 100%** và người dùng xác nhận đã nắm rõ 100% yêu cầu.
> 
> 2. **Bước 2: Phân tích tư duy chuyên sâu bằng BA & Solution Architecture**:
>    - Sau khi nắm rõ 100% yêu cầu, kích hoạt hội đồng kiến trúc sư giải pháp [`solution-architecture-council`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/solution-architecture-council/SKILL.md) và kỹ năng tư duy phản biện [`critical-thinking`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/critical-thinking/SKILL.md).
>    - Áp dụng First Principles, Root Cause Analysis (5 Whys), phân tích Trade-off đa chiều, Toulmin Model, kết hợp các kỹ năng BA/SA bổ trợ ([`assumption-validation`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/assumption-validation/SKILL.md), [`user-psychology`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/user-psychology/SKILL.md), [`business-fundamentals`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/business-fundamentals/SKILL.md), [`technical-basics`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/technical-basics/SKILL.md)).
> 
> 3. **Bước 3: Tinh chỉnh & Sáng tạo giải pháp Enterprise qua [`idea-refine`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/idea-refine/SKILL.md)**:
>    - Sau khi bàn luận tư duy ở Bước 2, sử dụng `idea-refine` để mở rộng góc nhìn mới (Divergent Thinking với 3-5 phương án khả dĩ).
>    - Stress-test và hội tụ (Convergent Thinking) chọn ra **phương pháp hợp lý, tối ưu, chuẩn Enterprise nhất** cho người dùng.
>    - Xuất bản One-Pager: Problem Statement, Recommended Direction (Enterprise), Key Assumptions, MVP Scope, và Not Doing List.
> 
> 4. **Bước 4: Tra cứu định tuyến skills qua [`using-agent-skills`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/using-agent-skills/SKILL.md) & Lập Plan chờ User duyệt trước khi thực thi**:
>    - Tra cứu Registry 53 Singleton Skills để chọn đúng các kỹ năng triển khai tiếp theo (UI Taste, TDD, DevTools, v.v.).
>    - Lập Implementation Plan chi tiết theo từng Vertical Slices kèm Definition of Done & Verification Checklist.
>    - **DỪNG LẠI (STOP & WAIT) CHỜ USER PHÊ DUYỆT PLAN TRƯỚC KHI THỰC THI BẤT KỲ DÒNG CODE NÀO**.

---

## Solution Architecture & Business Analysis Discovery Map

Khi tiếp nhận bất kỳ bài toán/yêu cầu nào, hãy tra cứu cây định tuyến dưới đây để kích hoạt đúng Singleton Skill:

```
Yêu Cầu Từ Người Dùng (User Request)
    │
    ▼
[BƯỚC 1: Cổng Phỏng Vấn 100% Tự Tin]
    └── Phỏng vấn 1-1 từng câu (Q & GUESS) cho tới khi CONFIDENCE = 100% ──→ interview-me
    │
    ▼
[BƯỚC 2: Phân Tích Chuyên Sâu BA & Solution Architecture Council]
    ├── Hội đồng SA Enterprise, Trade-offs, First Principles, STAR-P ───→ solution-architecture-council
    ├── Tư duy phản biện, 5 Whys Root Cause, Toulmin Argumentation ─────→ critical-thinking
    └── Phối hợp BA/SA skills bổ trợ:
        ├── Nhận diện & kiểm chứng giả định rủi ro ẩn ──────────────────→ assumption-validation
        ├── Giải mã tâm lý, rào cản hành vi của user ──────────────────→ user-psychology
        ├── Đánh giá mô hình nghiệp vụ, ROI & giá trị cốt lõi ──────────→ business-fundamentals
        └── Đánh giá tính khả thi kỹ thuật sơ bộ & trade-offs ──────────→ technical-basics
    │
    ▼
[BƯỚC 3: Tinh Chỉnh & Sáng Tạo Giải Pháp Enterprise]
    └── Divergent/Convergent Thinking, chọn phương án Enterprise tối ưu ─→ idea-refine
    │
    ▼
[BƯỚC 4: Định Tuyến Kỹ Năng & Lập Plan Chờ User Duyệt]
    ├── Tra cứu danh mục, điều phối kỹ năng thực thi ───────────────────→ using-agent-skills
    ├── Phân rã bài toán thành Vertical Slices có thứ tự phụ thuộc ────→ planning-and-task-breakdown
    └── 🛑 DỪNG LẠI CHỜ USER PHÊ DUYỆT PLAN TRƯỚC KHI CODE ─────────────→ (User Approval Gate)
        │
        └── KHI ĐÃ ĐƯỢC DUYỆT (Approved) ──► Kích hoạt các skills thực thi:
            │
            ├── 1. Requirements & Interface Design
            │   ├── Lập đặc tả kỹ thuật PRD, User Stories, A/C ─────────→ spec-driven-development
            │   ├── Phản biện rủi ro kiến trúc độc lập (Adversarial) ───→ doubt-driven-development
            │   ├── Thiết kế hợp đồng API & ranh giới module ───────────→ api-and-interface-design
            │   └── Thiết lập Semantic Design System (DESIGN.md) ───────→ stitch-design-taste
            │
            ├── 2. Frontend UI Taste & Engineering (Anti-Slop)
            │   ├── Xây dựng UI chuẩn production, Tailwind v4, a11y ────→ frontend-ui-engineering
            │   ├── Thiết kế UI anti-slop, chống generic template ──────→ design-taste-frontend
            │   ├── Thẩm mỹ visual cao cấp agency quốc tế (Awwwards) ───→ high-end-visual-design
            │   ├── Hoạt ảnh GSAP mượt mà & Bento grid hiện đại ────────→ gpt-taste
            │   └── Tái thiết kế nâng cấp dự án cũ lên chuẩn cao cấp ───→ redesign-existing-projects
            │
            ├── 3. Implementation & Verification
            │   ├── Triển khai mã nguồn theo Vertical Slices tuần tự ───→ incremental-implementation
            │   ├── Căn cứ mã nguồn trên tài liệu chính thống ──────────→ source-driven-development
            │   ├── Lập trình hướng kiểm thử (TDD - Test First) ────────→ test-driven-development
            │   ├── Kiểm thử runtime trên browser qua DevTools ─────────→ browser-testing-with-devtools
            │   ├── Review code đa trục trước khi merge (Gatekeeper) ───→ code-review-and-quality
            │   ├── Đối chiếu hội tụ 3 chiều & khắc phục sót việc ─────→ cross-artifact-convergence
            │   └── Ép buộc sinh code trọn vẹn 100%, cấm placeholder ───→ full-output-enforcement
            │
            └── 4. Operations, Hardening & Delivery
                ├── Khung tư duy STAR-P phân tích sự cố & RCA Playbook ─→ star-solution-architecture
                ├── Stop-the-line debugging & phục hồi lỗi ──────────────→ debugging-and-error-recovery
                ├── Phòng chống lỗ hổng bảo mật OWASP ──────────────────→ security-and-hardening
                ├── Tối ưu hóa hiệu năng & profiling bottlenecks ───────→ performance-optimization
                ├── Quản lý Git commit atomic, phân nhánh sạch ─────────→ git-workflow-and-versioning
                └── Tự động hóa CI/CD & Quality Gates ──────────────────→ ci-cd-and-automation
```

---

## Nguyên Tắc Vận Hành Singleton Skill

1. **One Role Per File (Nguyên lý Đơn Nhiệm)**: Mỗi file Skill trong `.agents/skills/` chỉ chịu trách nhiệm cho một miền kiến thức/quy trình duy nhất. Tuyệt đối không gộp chung các quy tắc không liên quan vào cùng một file.
2. **Khai Báo Skill Minh Bạch (Explicit Skill Citation)**: Trong mọi báo cáo phân tích, kế hoạch implementation hoặc walkthrough bàn giao, AI **BẮT BUỘC NÊU RÕ TÊN CỤ THỂ CỦA CÁC SKILLS ĐANG ĐƯỢC KÍCH HOẠT**.
3. **Tuân Thủ Thứ Tự Thực Thi (Strict Execution Order)**: Tuân thủ quy trình 4 bước bắt buộc: (1) `interview-me` $\rightarrow$ (2) BA/SA Council $\rightarrow$ (3) `idea-refine` $\rightarrow$ (4) `using-agent-skills` lên Plan chờ user duyệt.

---

## Danh Mục Tra Cứu Toàn Diện (53 Singleton Skills Catalog)

### 0. Meta-Orchestration & Core Council (3 Skills)
| Tên Skill | Trách Nhiệm Đơn Nhất (Singleton Responsibility) |
| :--- | :--- |
| [`using-agent-skills`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/using-agent-skills/SKILL.md) | Điều phối danh mục kỹ năng, bản đồ khám phá và quy tắc kích hoạt Singleton Skills. |
| [`solution-architecture-council`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/solution-architecture-council/SKILL.md) | Hội đồng Solution Architecture đa chuyên gia 50 năm kinh nghiệm, nguyên lý enterprise & khung STAR-P. |
| [`full-output-enforcement`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/full-output-enforcement/SKILL.md) | Ép buộc sinh code trọn vẹn 100%, cấm placeholder/cắt xén mã nguồn và chia tách token an toàn. |

### 1. Requirements, Inception & Business Analysis (18 Skills)
| Tên Skill | Phân Loại | Trách Nhiệm Đơn Nhất (Singleton Responsibility) |
| :--- | :--- | :--- |
| [`interview-me`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/interview-me/SKILL.md) | Inception Core | Phỏng vấn tương tác 1-1 từng câu hỏi để bóc tách intent thực sự của người dùng đến khi tự tin đạt 100%. |
| [`idea-refine`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/idea-refine/SKILL.md) | Inception Core | Tinh chỉnh ý tưởng mới, tư duy phân kỳ/hội tụ, chọn phương pháp tối ưu enterprise nhất, chốt MVP Scope One-Pager. |
| [`spec-driven-development`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/spec-driven-development/SKILL.md) | Inception Core | Xây dựng đặc tả kỹ thuật PRD, User Stories, Acceptance Criteria trước khi lập trình. |
| [`doubt-driven-development`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/doubt-driven-development/SKILL.md) | Inception Core | Phản biện rủi ro kiến trúc độc lập (Adversarial Review) đối với các quyết định kỹ thuật quan trọng. |
| [`brandkit`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/brandkit/SKILL.md) | Inception Core | Thiết lập bộ nhận diện thương hiệu, hệ thống logo, identity deck và typography cao cấp. |
| [`critical-thinking`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/critical-thinking/SKILL.md) | Business Analysis | Tư duy phản biện logic, phân tách bài toán đa tầng, kiểm chứng luận cứ và truy tìm nguyên nhân gốc. |
| [`assumption-validation`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/assumption-validation/SKILL.md) | Business Analysis | Phát hiện các giả định rủi ro ẩn giấu, kiểm tra tính đúng đắn và phát hiện mâu thuẫn yêu cầu trước khi làm. |
| [`insight-extraction`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/insight-extraction/SKILL.md) | Business Analysis | Đào sâu dữ liệu/yêu cầu tìm ra cơ hội đột phá, nguyên nhân cốt lõi và các hệ quả phi hiển ngôn. |
| [`pattern-recognition`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/pattern-recognition/SKILL.md) | Business Analysis | Nhận diện mô thức lặp lại, quy luật hành vi, xu hướng và các điểm bất thường (outliers). |
| [`data-analysis`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/data-analysis/SKILL.md) | Business Analysis | Phân tích dữ liệu thực tế, lượng hóa bài toán, kiểm chứng giả định bằng số liệu và bằng chứng. |
| [`doc-synthesis`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/doc-synthesis/SKILL.md) | Business Analysis | Tổng hợp tài liệu phức tạp, trích xuất thông điệp cốt lõi thành bản Executive Summary sắc bén. |
| [`communication`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/communication/SKILL.md) | Business Analysis | Chuẩn hóa tài liệu truyền thông, bảo vệ quan điểm giải pháp, thuyết phục và diễn đạt logic theo người nghe. |
| [`business-fundamentals`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/business-fundamentals/SKILL.md) | Business Analysis | Đánh giá mô hình kinh doanh, unit economics, cơ chế sinh giá trị, ROI và mức độ ưu tiên tính năng. |
| [`market-dynamics`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/market-dynamics/SKILL.md) | Business Analysis | Đánh giá bối cảnh thị trường, định vị giải pháp, phân tích đối thủ cạnh tranh và phân khúc người dùng. |
| [`user-psychology`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/user-psychology/SKILL.md) | Business Analysis | Giải mã tâm lý học hành vi, thiên kiến nhận thức, rào cản thích ứng và động lực sử dụng của người học. |
| [`stakeholder-management`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/stakeholder-management/SKILL.md) | Business Analysis | Quản trị kỳ vọng của các bên liên quan, giải quyết xung đột lợi ích và xây dựng đồng thuận giải pháp. |
| [`compliance-risk`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/compliance-risk/SKILL.md) | Business Analysis | Nhận diện rủi ro tuân thủ quy định, quyền riêng tư dữ liệu, ràng buộc an toàn và tính liên tục nghiệp vụ. |
| [`technical-basics`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/technical-basics/SKILL.md) | Business Analysis | Đánh giá tính khả thi kỹ thuật sơ bộ, trade-offs kiến trúc và chuyển ngữ yêu cầu nghiệp vụ sang kỹ thuật. |

### 2. Architecture & Design (5 Skills)
| Tên Skill | Trách Nhiệm Đơn Nhất (Singleton Responsibility) |
| :--- | :--- |
| [`planning-and-task-breakdown`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/planning-and-task-breakdown/SKILL.md) | Phân rã bài toán thành các task kỹ thuật có thứ tự, cô lập phạm vi và kiểm chứng được. |
| [`api-and-interface-design`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/api-and-interface-design/SKILL.md) | Thiết kế hợp đồng API, schema Zod và ranh giới module bất biến giữa Client và Server. |
| [`documentation-and-adrs`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/documentation-and-adrs/SKILL.md) | Ghi nhận các quyết định kiến trúc quan trọng (Architectural Decision Records) và tài liệu kỹ thuật. |
| [`deprecation-and-migration`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/deprecation-and-migration/SKILL.md) | Chiến lược di trú hệ thống cũ, quản lý deprecation và chuyển đổi an toàn không gián đoạn. |
| [`stitch-design-taste`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/stitch-design-taste/SKILL.md) | Thiết lập Semantic Design System (`DESIGN.md`), quy chuẩn spacing, palette màu và layout. |

### 3. High-Quality Implementation & Visual Engineering (14 Skills)
| Tên Skill | Trách Nhiệm Đơn Nhất (Singleton Responsibility) |
| :--- | :--- |
| [`incremental-implementation`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/incremental-implementation/SKILL.md) | Triển khai mã nguồn theo từng lát cắt dọc (Vertical Slices), chuyển giao từng bước kiểm chứng được. |
| [`source-driven-development`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/source-driven-development/SKILL.md) | Căn cứ mã nguồn trên tài liệu chính thống của framework/thư viện, loại bỏ pattern lỗi thời. |
| [`context-engineering`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/context-engineering/SKILL.md) | Tối ưu ngữ cảnh làm việc cho AI, quản lý luật lệ và cấu hình prompt chuẩn mực. |
| [`frontend-ui-engineering`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/frontend-ui-engineering/SKILL.md) | Xây dựng giao diện chất lượng cao chuẩn production, tối ưu a11y, state management và CSS. |
| [`image-to-code`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/image-to-code/SKILL.md) | Hiện thực hóa mã nguồn frontend pixel-perfect từ ảnh mockup thiết kế. |
| [`design-taste-frontend`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/design-taste-frontend/SKILL.md) | Thiết kế giao diện anti-slop, chống template generic rẻ tiền cho landing pages và portfolios. |
| [`design-taste-frontend-v1`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/design-taste-frontend-v1/SKILL.md) | Bản v1 bảo tồn tương thích ngược của design-taste. |
| [`high-end-visual-design`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/high-end-visual-design/SKILL.md) | Chuẩn thiết kế visual cao cấp như agency quốc tế hàng đầu (Awwwards-tier). |
| [`gpt-taste`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/gpt-taste/SKILL.md) | Kỹ sư hoạt ảnh UX/UI & GSAP mượt mà, cấu trúc AIDA và bento grid hiện đại. |
| [`minimalist-ui`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/minimalist-ui/SKILL.md) | Giao diện tối giản biên tập (Editorial Minimalism), monochrome ấm, flat grid thanh lịch. |
| [`industrial-brutalist-ui`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/industrial-brutalist-ui/SKILL.md) | Giao diện Industrial Brutalism & Tactical Telemetry cho dashboard, terminal và blueprints. |
| [`redesign-existing-projects`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/redesign-existing-projects/SKILL.md) | Tái thiết kế nâng cấp dự án cũ lên chuẩn cao cấp mà không làm gãy logic chức năng. |
| [`imagegen-frontend-web`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/imagegen-frontend-web/SKILL.md) | Chỉ đạo hình ảnh và khởi tạo concept UI riêng biệt từng section cho website. |
| [`imagegen-frontend-mobile`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/imagegen-frontend-mobile/SKILL.md) | Khởi tạo concept màn hình app mobile native đặt trong mockup điện thoại sắc nét. |

### 4. Diagnostics, Problem-Solving & Verification (10 Skills)
| Tên Skill | Trách Nhiệm Đơn Nhất (Singleton Responsibility) |
| :--- | :--- |
| [`risk-management`](file:///home/huyhg/Documents/VBA_Consult/MetadataCompiler/.agents/skills/risk-management/SKILL.md) | **⚡ KÍCH HOẠT TRƯỚC MỌI CODE BLOCK**: Phân tích 5 trục rủi ro (User Behavior, Data Integrity, System Env, Pipeline State, Resource Leak), bắt buộc áp dụng Guard Clause + safe_* helpers + Cleanup pattern + Transaction Rollback. |
| [`star-solution-architecture`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/star-solution-architecture/SKILL.md) | Khung tư duy STAR-P phân tích sự cố, tìm nguyên nhân gốc rễ và tích lũy Principle Playbook. |
| [`debugging-and-error-recovery`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/debugging-and-error-recovery/SKILL.md) | Quy trình gỡ lỗi có cấu trúc Stop-the-line, tái hiện lỗi và khắc phục tận gốc. |
| [`test-driven-development`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/test-driven-development/SKILL.md) | Lập trình hướng kiểm thử (TDD), viết test thất bại trước khi viết code triển khai. |
| [`browser-testing-with-devtools`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/browser-testing-with-devtools/SKILL.md) | Kiểm thử runtime trên trình duyệt qua DevTools (DOM, Network, Console errors). |
| [`code-review-and-quality`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/code-review-and-quality/SKILL.md) | Rà soát chất lượng đa trục (Reviewer Gate) trước khi merge hoặc chuyển giao QA. |
| [`cross-artifact-convergence`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/cross-artifact-convergence/SKILL.md) | Đối chiếu hội tụ 3 chiều (Spec <-> Plan <-> Tasks <-> Code), phát hiện việc sót và sinh remediation tasks. |
| [`security-and-hardening`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/security-and-hardening/SKILL.md) | Phòng chống lỗ hổng bảo mật OWASP, kiểm soát quyền truy cập và xác thực đầu vào. |
| [`performance-optimization`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/performance-optimization/SKILL.md) | Đo lường hiệu năng trước khi tối ưu, profiling và xử lý triệt để điểm nghẽn (bottlenecks). |
| [`code-simplification`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/code-simplification/SKILL.md) | Tinh giản code, giảm độ phức tạp nhận thức (cyclomatic complexity) mà không đổi hành vi. |

### 5. Operations, CI/CD & Delivery (4 Skills)
| Tên Skill | Trách Nhiệm Đơn Nhất (Singleton Responsibility) |
| :--- | :--- |
| [`git-workflow-and-versioning`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/git-workflow-and-versioning/SKILL.md) | Quản trị Git commit atomic, phân nhánh sạch và chiến lược hợp nhất code. |
| [`ci-cd-and-automation`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/ci-cd-and-automation/SKILL.md) | Thiết lập pipeline CI/CD tự động hóa, kiểm soát chất lượng qua quality gates. |
| [`observability-and-instrumentation`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/observability-and-instrumentation/SKILL.md) | Thiết lập logging có cấu trúc, RED metrics, tracing và cơ chế cảnh báo vận hành. |
| [`shipping-and-launch`](file:///home/huyhg/Documents/ME-RESEARCH/.agents/skills/shipping-and-launch/SKILL.md) | Checklist tiền phát hành, chiến lược phát hành an toàn (staged rollout) và kế hoạch rollback. |
