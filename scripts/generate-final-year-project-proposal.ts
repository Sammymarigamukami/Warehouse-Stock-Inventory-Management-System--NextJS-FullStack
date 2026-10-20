import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import {
  AlignmentType,
  Document,
  Footer,
  HeadingLevel,
  Packer,
  PageNumber,
  Paragraph,
  TabStopPosition,
  TabStopType,
  TextRun,
} from "docx";

const outputPath = path.resolve(process.cwd(), "docs", "final-year-project-proposal.docx");

const bodyFont = "Calibri";
const headingFont = "Calibri";

const chapter = (title: string) =>
  new Paragraph({
    text: title,
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 200 },
    thematicBreak: false,
  });

const sectionHeading = (title: string) =>
  new Paragraph({
    text: title,
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 120 },
  });

const body = (text: string) =>
  new Paragraph({
    spacing: { after: 120, line: 360 },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text, font: bodyFont, size: 24 })],
  });

async function generateProposalDocx() {
  const titlePage: Paragraph[] = [
    new Paragraph({
      children: [
        new TextRun({
          text: "FINAL-YEAR PROJECT PROPOSAL",
          bold: true,
          size: 36,
          font: headingFont,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 2400, after: 300 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Warehouse Stock Inventory Management System",
          bold: true,
          size: 30,
          font: headingFont,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 600 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "A Full-Stack Proposal Based on Next.js 16, React 19, TypeScript, Prisma, and MongoDB",
          italics: true,
          size: 24,
          font: bodyFont,
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 1200 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Prepared for Final-Year Project Assessment", size: 24, font: bodyFont })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 180 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Repository Context: Sammymarigamukami/Warehouse-Stock-Inventory-Management-System--NextJS-FullStack", size: 22, font: bodyFont })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 180 },
    }),
    new Paragraph({
      children: [new TextRun({ text: `Date: ${new Date().toISOString().slice(0, 10)}`, size: 22, font: bodyFont })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 600 },
      pageBreakBefore: false,
    }),
  ];

  const content: Paragraph[] = [
    ...titlePage,
    chapter("CHAPTER ONE: INTRODUCTION AND BACKGROUND"),
    body("This proposal presents the planned development of a Warehouse Stock Inventory Management System designed to solve operational inefficiencies in inventory tracking, procurement coordination, order fulfillment, and financial reconciliation for modern small-to-medium enterprises. The project is grounded in a production-style full-stack codebase built with Next.js 16, React 19, TypeScript, Prisma, MongoDB, TanStack Query, JWT-based authentication, and shadcn/ui with Tailwind CSS, with integrated services for payment, logistics, email communication, media handling, queuing, and observability."),
    body("Warehousing and stock governance remain a strategic challenge because many organizations still rely on fragmented spreadsheets, disconnected billing workflows, and delayed stock visibility. These limitations introduce avoidable stock-outs, overstocking, delayed dispatch, and weak traceability of business events. The proposed system addresses this gap by centralizing products, warehouses, suppliers, clients, orders, invoices, support tickets, and analytics in one role-aware digital platform."),
    sectionHeading("1.1 Problem Statement"),
    body("Many inventory systems used by growing businesses do not provide real-time coordination between stock movement, order lifecycle, invoice lifecycle, payment status, and role-based decision support. As a result, teams face conflicting data versions, weak accountability trails, and delayed responses to low-stock or payment exceptions. The core problem is the absence of a secure, integrated, and analytics-enabled system that links operational transactions to managerial insight."),
    sectionHeading("1.2 Aim"),
    body("To design and implement a robust, secure, and scalable Warehouse Stock Inventory Management System that unifies warehouse operations, product catalog management, order and invoice processing, role-based access control, and decision-support analytics within a single full-stack web platform."),
    sectionHeading("1.3 Objectives"),
    body("The project objectives are to: (i) design centralized modules for products, categories, suppliers, warehouses, clients, orders, invoices, and support workflows; (ii) implement role-sensitive access and workflow control for administrators, suppliers, clients, and staff; (iii) support payment and fulfillment coordination through integrations such as Stripe and Shippo; (iv) provide reliable notifications and messaging channels through Brevo and queue-backed processing with Redis/QStash; (v) incorporate monitoring, validation, and error tracking through Sentry and structured server-side safeguards; and (vi) deliver analytical dashboards for operational visibility and strategic planning."),
    sectionHeading("1.4 Research Questions"),
    body("This proposal is guided by the following questions: How can a Next.js-based full-stack architecture provide coherent synchronization between stock states and commercial transactions? Which security and validation controls are necessary to protect warehouse and financial workflows in role-diverse environments? What architectural and caching strategies improve data freshness, responsiveness, and reliability for inventory-intensive user journeys? How can integrated analytics improve planning decisions on stock, suppliers, and order performance?"),
    sectionHeading("1.5 Significance of the Study"),
    body("Academically, the project contributes a practical model for full-stack enterprise application engineering that combines frontend responsiveness, backend integrity controls, and operational observability. Practically, it offers organizations a pathway to reduce manual reconciliation workload, improve stock allocation accuracy, support safer financial processing, and increase service quality through faster issue resolution and clearer audit trails."),
    sectionHeading("1.6 Scope and Delimitations"),
    body("The proposal scope covers web-based management of inventory and related workflows: catalog management, warehouse allocation, order and invoice lifecycle handling, payment status coordination, shipment workflows, support ticketing, and dashboard analytics. The work includes role-based interfaces, API validation, error monitoring, and integration points with third-party services. It does not include native mobile application development, direct hardware integration for IoT sensors, or large-scale machine-learning model training beyond practical forecasting features supported by the existing application context."),
    sectionHeading("1.7 Limitations and Assumptions"),
    body("Potential limitations include dependence on external service reliability (for example payment gateways and shipping providers), internet availability for cloud-hosted operations, and configuration complexity across environments. The proposal assumes the availability of required environment variables, secure deployment practices, and representative test data for evaluating workflow correctness. The project frames outcomes as expected deliverables rather than claiming completed empirical performance gains."),

    chapter("CHAPTER TWO: LITERATURE REVIEW"),
    body("The literature on inventory information systems consistently emphasizes integration, real-time visibility, data integrity, and decision support as core success factors. Earlier systems often separated stock records from billing and supplier communication, causing latency and inconsistency. Contemporary practice trends toward unified full-stack platforms where transactional modules and analytics coexist, reducing synchronization overhead and improving process accountability."),
    sectionHeading("2.1 Inventory System Evolution and Architectural Trends"),
    body("Research and industry implementations show a migration from monolithic desktop tools and spreadsheet workflows toward modular web platforms with API-driven services. The current project aligns with this trajectory by using Next.js and TypeScript for cohesive client-server development while preserving separation of concerns across server logic, route handlers, and reusable UI components."),
    sectionHeading("2.2 Data Consistency, Caching, and Synchronization"),
    body("A central theme in the literature is balancing responsiveness with data integrity. Systems employing client-side caching must prevent stale reads after critical write operations. The project context addresses this through explicit query invalidation patterns with TanStack Query, targeted cache key strategies, and server-side data validation. Comparable approaches in modern ERP-lite systems similarly combine optimistic UI behavior with authoritative server reconciliation."),
    sectionHeading("2.3 Security and Access Control"),
    body("Studies on enterprise application security stress layered controls: authentication, authorization, input validation, least privilege, and auditable error handling. The proposed system reflects these principles through JWT authentication, role-based access pathways, Zod validation of user-facing payloads, and selective logging/monitoring workflows. These mechanisms are consistent with recommended secure-by-default design practices for multi-actor systems handling commercial and identity-sensitive data."),
    sectionHeading("2.4 Integrated Commerce and Logistics Workflows"),
    body("Prior work indicates that business value increases when inventory systems are tightly linked to payment and shipping channels. Integrated payment reconciliation reduces invoice ambiguity, while shipping status feeds improve fulfillment transparency. The project’s architecture, including Stripe and Shippo integration points, follows this approach by connecting order and invoice transitions with fulfillment events and status synchronization."),
    sectionHeading("2.5 Observability, Reliability, and Support Operations"),
    body("Reliable platforms require proactive monitoring and support visibility. Literature on DevOps-informed information systems highlights tracing, structured error capture, and user-support channels as essential for maintainability. The proposed use of Sentry monitoring, server-side error-response helpers, and support ticket management reflects this perspective by enabling rapid issue triage and operational learning loops."),
    sectionHeading("2.6 Gap Analysis and Positioning"),
    body("Many documented systems treat inventory, payment, shipping, support, and analytics as loosely coupled modules with manual bridging. This proposal positions the Warehouse Stock Inventory Management System as a coherent integration-first platform that embeds these capabilities into a role-aware user experience. The expected contribution is a practical blueprint for scalable, maintainable, and academically rigorous full-stack inventory system development."),

    chapter("CHAPTER THREE: METHODOLOGY"),
    body("The project will adopt an iterative software engineering methodology that combines requirements-driven planning with incremental development, continuous validation, and feedback-led refinement. The methodology is suitable for complex systems where multiple modules and integrations must be coordinated while preserving quality and maintainability."),
    sectionHeading("3.1 Research and Development Approach"),
    body("A design-and-build methodology will be used, with Agile-inspired iteration cycles: requirement clarification, architecture refinement, module implementation, integration, testing, and review. Each cycle focuses on delivering a verifiable increment, such as stock allocation flow hardening, invoice payment synchronization, or role-based detail-page access controls."),
    sectionHeading("3.2 Data Collection and Requirement Elicitation"),
    body("Requirements will be derived from domain expectations of warehouse operations and from observable repository capabilities, including entities, route structures, validation patterns, and UI workflows. Functional expectations are mapped from existing module coverage: products, categories, suppliers, warehouses, clients, orders, invoices, analytics, authentication, notifications, and support tickets. Non-functional expectations are inferred from monitoring, caching, error handling, and testing practices already present in the system context."),
    sectionHeading("3.3 Development Stack Justification"),
    body("Next.js 16 and React 19 provide a modern full-stack web framework with strong routing and rendering support. TypeScript enforces static typing for safer refactoring and domain modeling. Prisma with MongoDB supports structured data access for complex business entities. TanStack Query enables controlled client-state synchronization. JWT authentication and role checks enforce access boundaries. Tailwind and shadcn/ui accelerate consistent component-driven UI implementation. Integrated services (Stripe, Shippo, Brevo, ImageKit, Redis/QStash, Sentry) provide extensible operational capability."),
    sectionHeading("3.4 Implementation Procedure"),
    body("Implementation will proceed in phases: (i) refine domain models and API contracts; (ii) complete and validate core CRUD and lifecycle workflows; (iii) integrate payment, shipping, messaging, and media pipelines; (iv) enforce role-based guards and validation constraints; (v) optimize cache invalidation and SSR-client synchronization for responsive UX; and (vi) harden observability and support tooling. Each phase will include peer review and regression checks before progression."),
    sectionHeading("3.5 Testing and Evaluation Strategy"),
    body("Evaluation will combine automated and scenario-based testing. Automated checks include linting, TypeScript build validation, and unit/integration tests through Vitest. Workflow-oriented checks will verify key interactions such as stock reservation/release, order-to-invoice linkage, partial/full payment transitions, shipment status updates, and role-restricted operations. Reliability indicators will include error trace quality, cache coherence under mutation-heavy flows, and consistency of dashboard figures with transactional records."),
    sectionHeading("3.6 Ethical and Quality Considerations"),
    body("The project will follow secure coding and privacy-conscious design principles. Sensitive credentials remain environment-driven, role boundaries are enforced, and error reporting avoids exposing private data. Quality assurance will prioritize correctness, maintainability, and transparent documentation of assumptions and limitations."),

    chapter("CHAPTER FOUR: SYSTEM ANALYSIS"),
    body("System analysis defines the business context, user actors, functional services, and constraints that shape implementation decisions for the Warehouse Stock Inventory Management System."),
    sectionHeading("4.1 Stakeholders and System Actors"),
    body("Primary actors include administrators, store owners, suppliers, clients, and operational staff. Administrators require global visibility, configuration control, and user oversight. Suppliers need scoped access to product, order, and invoice context relevant to their participation. Clients require transparent purchasing and invoice tracking views. Staff need efficient transaction handling and support-response workflows."),
    sectionHeading("4.2 Existing Process Challenges"),
    body("Common pre-system pain points include delayed stock visibility across warehouses, manual order-to-invoice reconciliation, fragmented communication with suppliers and clients, inconsistent exception handling, and weak observability for production issues. These weaknesses increase transaction risk and reduce confidence in operational reporting."),
    sectionHeading("4.3 Functional Requirements"),
    body("The system shall provide: secure authentication and role-based authorization; management of products, categories, suppliers, and warehouses; stock allocation and transfer workflows; order creation, update, and fulfillment lifecycle handling; invoice generation, status progression, and payment tracking; shipment workflow linkage; dashboard analytics for key KPIs; notification and support ticket workflows; media handling for product assets and QR-code support; and audit-friendly activity visibility."),
    sectionHeading("4.4 Non-Functional Requirements"),
    body("The system should meet non-functional requirements for performance, usability, reliability, security, maintainability, and scalability. It should provide responsive interaction through optimized caching and SSR hydration patterns, robust error handling through monitoring integrations, strong input validation and access checks, and modular architecture that supports controlled future expansion."),
    sectionHeading("4.5 Feasibility Analysis"),
    body("Technical feasibility is high due to the mature stack and existing module foundations. Operational feasibility is supported by browser-based access and role-oriented interfaces that fit common business workflows. Economic feasibility is enhanced by open-source core technologies and selective use of external services. Schedule feasibility is improved by iterative delivery and test-backed verification practices."),
    sectionHeading("4.6 Workflow and Data Flow Overview"),
    body("Core workflows include: product onboarding and categorization; warehouse stock assignment and reallocation; order composition with stock checks; invoice issuance linked to order state; payment progression handling including partial settlements; shipping state updates; and post-transaction support interactions. Data flows across API routes, domain repositories, cache synchronization layers, and user-facing dashboards, with monitoring signals captured for reliability oversight."),

    chapter("CHAPTER FIVE: SYSTEM DESIGN AND IMPLEMENTATION"),
    body("This chapter translates analysis outputs into concrete architecture, data design, security controls, integration plans, and implementation strategy aligned with the repository’s established technology context."),
    sectionHeading("5.1 Architectural Design"),
    body("The proposed architecture follows a full-stack web pattern where Next.js App Router supports server and client concerns within a cohesive TypeScript codebase. Server-side route handlers coordinate domain operations and response policies, while React components provide role-sensitive user interfaces. TanStack Query manages server-state synchronization and mutation invalidation for responsive interactions. This design supports modular growth while minimizing cross-layer drift."),
    sectionHeading("5.2 Database and Domain Model Design"),
    body("The data model centers on core entities: users, products, categories, suppliers, warehouses, stock allocations, orders, invoices, payments, shipment records, notifications, and support tickets. Prisma acts as the data access layer to MongoDB, enabling typed schema evolution and consistent query patterns. Design emphasis is placed on preserving relationship integrity, lifecycle timestamps, and status fields required for auditing and analytics."),
    sectionHeading("5.3 Security and Access Design"),
    body("Security design includes JWT-based session authentication, role-aware route gating, input validation for API payloads, and monitored error pipelines. Access policies are applied to protect entity details and mutation rights by actor role. The design also incorporates controlled logging behavior to avoid unnecessary exposure of expected client errors while escalating true server failures for diagnostics."),
    sectionHeading("5.4 Integration Design"),
    body("Integration interfaces are proposed for payment processing (Stripe), shipping operations (Shippo), email and notifications (Brevo and queued job processing through Redis/QStash), media management (ImageKit), and observability (Sentry). The implementation strategy treats external service calls as bounded dependencies with explicit failure handling, retriable workflows where appropriate, and status reconciliation to keep internal records authoritative."),
    sectionHeading("5.5 User Interface and Experience Design"),
    body("The UI design uses shadcn/ui primitives and Tailwind-based styling for consistency, accessibility, and maintainability. Role-based views expose only relevant actions while preserving clarity of status indicators, summary cards, and detail workflows. The interface emphasizes dense but readable operational information across products, orders, invoices, warehouses, analytics, and support modules."),
    sectionHeading("5.6 Implementation Plan"),
    body("Implementation will be organized into iterative milestones: architecture and schema stabilization, core transactional module completion, integration hardening, analytics completion, and production-readiness validation. Each milestone includes linting, build verification, test execution, and quality gates for regression control. Deployment readiness will include environment validation, observability checks, and post-release monitoring routines."),
    sectionHeading("5.7 Testing, Verification, and Evaluation Criteria"),
    body("Verification will assess correctness of business rules, consistency of cross-module data updates, and reliability under realistic mutation patterns. Evaluation criteria include successful execution of automated tests, stable build outcomes, integrity of role-based permissions, correctness of payment and shipment state transitions, and expected dashboard coherence with transactional records."),
    sectionHeading("5.8 Expected Outcomes"),
    body("Expected project outcomes include a deployable and maintainable full-stack warehouse inventory platform, reduced manual reconciliation effort, improved operational visibility, stronger role-governed data protection, and a reusable engineering blueprint for future inventory-centric systems. These outcomes are proposed targets aligned with the repository’s implementation direction rather than claims of already measured field performance."),

    chapter("CHAPTER SIX: CONCLUSION AND RECOMMENDATIONS"),
    body("Note on chapter numbering: the user-supplied chapter list repeated Chapter Four and Chapter Five labels; numbering has been normalized sequentially in this proposal to preserve structure clarity."),
    sectionHeading("6.1 Conclusion"),
    body("This proposal defines a technically feasible and academically grounded plan for building a comprehensive Warehouse Stock Inventory Management System using a modern TypeScript full-stack architecture. The project addresses practical inventory governance challenges by integrating stock management, commercial workflows, role-based access, and observability within one cohesive platform. The methodology, analysis, and design sections establish a clear pathway from requirement understanding to implementation and quality assurance."),
    sectionHeading("6.2 Recommendations"),
    body("It is recommended that implementation proceed with strict requirement traceability and phased delivery to preserve system stability. Security and validation controls should remain first-class priorities, especially around payment, authentication, and role-sensitive resources. Continuous monitoring and cache-coherence testing should be maintained during development and deployment. Future enhancements can include deeper predictive analytics, broader external ERP interoperability, and expanded reporting automation once the core transactional reliability baseline is fully validated."),
  ];

  const doc = new Document({
    styles: {
      default: {
        heading1: {
          run: {
            bold: true,
            font: headingFont,
            size: 30,
            color: "1F2937",
          },
          paragraph: {
            spacing: { after: 180 },
          },
        },
        heading2: {
          run: {
            bold: true,
            font: headingFont,
            size: 26,
            color: "334155",
          },
          paragraph: {
            spacing: { after: 100 },
          },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440,
            },
          },
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
                children: [
                  new TextRun({ text: "Page ", size: 20, font: bodyFont }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 20, font: bodyFont }),
                  new TextRun({ text: " of ", size: 20, font: bodyFont }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 20, font: bodyFont }),
                ],
                alignment: AlignmentType.RIGHT,
              }),
            ],
          }),
        },
        children: content,
      },
    ],
  });

  await mkdir(path.dirname(outputPath), { recursive: true });
  const buffer = await Packer.toBuffer(doc);
  await writeFile(outputPath, buffer);

  console.log(`Generated proposal document at: ${outputPath}`);
}

generateProposalDocx().catch((error) => {
  console.error("Failed to generate proposal document:", error);
  process.exitCode = 1;
});
