# SAD-Template.pdf — Structure Reference

Extracted from the official SBD `SAD-Template.pdf` (5 pages). Use this as the authoritative
section list and guidance when drafting a Solution Architecture Document. Every generated SAD
must include all headings below, in this order, with real project content (no placeholder text
left in the final document).

## 1 Solution Design Summary

### 1.1 Solution Design Overview
Narrative overview of what the solution covers and why.

### 1.2 Solution Design Constraints
Highlight constraints that resulted in exceptions or trade-offs. In particular:
- Time scale challenges — e.g. why the solution will not be delivered on-time.
- Legacy solutions — e.g. if the solution leverages sub-optimal components and why it can't be
  redesigned within the current project (risk, cost, time, commercials, etc.).

### 1.3 Solution Design Risks / Contingencies
For all constraints, out-of-tolerance design aspects, and risks, highlight:
- How each is mitigated (i.e. won't cause unacceptable impact).
- Future plans to remove such items from the solution.

## 2 Solution Design

### 2.1 Solution Architecture Diagram
- Provide a Collapsed Level 2 Solution Architecture Design Diagram (remove reference labels for
  the screenshot) — normally copy/pasted from Ardoq.
- Provide an Expanded Level 2 Solution Architecture Design Diagram — normally copy/pasted from
  Ardoq.
- If Ardoq access/export is unavailable, generate a clearly-labeled conceptual substitute diagram
  (current state + future state) and note that it should be replaced with the Ardoq export prior
  to SARB/TechPASS submission.

### 2.2 Process Sequence Diagrams
**Critical section — must be diligently completed.** For each high-level requirement from the
BRD, create a Sequence Diagram illustrating how the architecture will meet that business process.
Provide an image for each diagram; if a clean rendered diagram isn't achievable, provide (or
supplement with) an editable script for websequencediagrams.com.

### 2.3 Solution Design Details

#### 2.3.1 Systems
All systems involved in the solution, not just large/prominent ones — include legacy systems,
satellite systems, manually connected systems, reporting solutions, and middleware. (Definitive
listing belongs in section 2.4.)

#### 2.3.2 System Connections
The system connections defining system linkage, including method (e.g. Direct, Middleware,
consolidated via a Data Lake, etc.).

#### 2.3.3 Integration Approaches
Summarize desired integration approaches (APIs, flat files, DB log replication, etc.). May change
during design phase, with all architecture changes approved by SA. (Specific integration details
belong in section 2.4.)

#### 2.3.4 Middleware & Transfer Technologies
Summarize all middleware/transfer technology used (e.g. Agora, SFTP, Direct Integration, etc.).
(Specific details belong in section 2.4.)

#### 2.3.5 Data Objects
Data objects within scope, including which systems are the "master" for each object (e.g.
Customer, Material, Employee, Warranty, Inventory). Specific tables/fields/mappings are not
required here. (Interface-level detail belongs in section 2.4.)

#### 2.3.6 Data Strategy
Define the data strategy per object — periodic vs. immediate (Live/Batch), Full Extracts vs.
Changes only (Full/Delta). **SA standard: all integrations should be Delta, Asynchronous,
Real-Time interfaces.** (Per-integration detail belongs in section 2.4.)

#### 2.3.7 System Hosting
Where systems are hosted (SBD or External Network) and hosting technology (On-prem, SaaS, PaaS,
IaaS), why that approach was selected, and how it positively/negatively impacts the design.
Include a discussion of potential security impacts.

#### 2.3.8 System Existence
State clearly which systems are Existing vs. New To Be Implemented, and the reasoning (why this
approach is beneficial/correct for the project).

#### 2.3.9 Deployment & Support
Highlight which systems are expected to be Supported or Delivered by which party (SBD vs. Delivery
Vendor, MSP, Technology Vendor, etc.). Not a commercial document, but must be agreed with the
Build Lead.

### 2.4 Solution Design Interface Inventory
Complete listing of all integrations/interfaces for project reference — normally extracted from
the organization's approved Ardoq report and pasted in tabular/text format. Note where the Ardoq
master list is maintained if the live extract isn't available.
