# TravAll -- HCI Design Project (Full UX Lifecycle)

> **Human-Computer Interaction Design** | Figma + Python + User Research

A 6-phase HCI design project for **TravAll**, a travel management super-app that consolidates passports, visas, trip planning, and community insights into a single platform. Team project (4 members) completed over a full semester.

## Problem

- 58% of travelers face visa uncertainty and confusion
- 70% suffer from document clutter spread across 3-4 separate apps
- No unified platform exists for managing travel documents, trip planning, and visa requirements

## Design Process

### Phase 1: System Design Thinking
Design Thinking methodology with empathy mapping to define the problem space.

### Phase 2: User Research
Collected **94+ survey responses** analyzing traveler pain points, behaviors, and demographics.

### Phase 3: Personas and Storyboards
Three detailed personas derived from research data:
- Efficient minimalist traveler (chemistry student)
- High-frequency business voyager (12+ trips/year)
- Anxious first-timer with visa anxiety

Narrative storyboards showing before/after scenarios for each persona.

### Phase 4: Task Analysis and Information Architecture
- Hierarchical Task Analysis (HTA) for 3 key user scenarios
- System map showing screen relationships and navigation flows
- 30+ wireframes covering all core screens

### Phase 5A / 6A: Lo-Fi Prototype and Field Trial
- 17-screen clickable prototype built in Figma, then converted to an interactive HTML/JS browser prototype
- Field trial with 5 participants across 15 sequential tasks
- **96% task completion rate**, average completion time 6'54"
- **8.6/10 likability**, **9.0/10 visual clarity**
- Identified 3 specific usability problems (settings discoverability, button similarity, affordance issues)

### Phase 5B / 6B: High-Fidelity Prototype and Iteration
- Systematic fixes for all 3 identified usability problems
- Added notification bell (4 taps reduced to 1), redesigned button hierarchy, improved affordance cues
- New screens: Privacy/Security, Appearance (dark mode), Help/Support

## Automation Scripts

Python scripts for automating deliverable generation (~1,000 lines total):

| Script | Purpose |
|--------|---------|
| `hta_diagrams.py` | Generates publication-quality HTA diagrams via Graphviz |
| `create_presentation.py` | Builds Phase 3 PPTX from PDF storyboard extraction |
| `create_presentation_phase5.py` | Phase 5 presentation generator |
| `scripts/generate_phase4.py` | Phase 4 report generator with system maps |
| `scripts/generate_pptx.py` | Phase 4 presentation builder |
| `check_ai_score.py` | AI detection scoring for academic integrity |

## Project Structure

```
travall-hcid-project/
├── README.md
├── IE48L A1.pdf                    # Phase 1: System Design Thinking
├── IE48L A2.pdf                    # Phase 2: User Research (94+ responses)
├── IE48L_Assignment2.md            # HTA analysis (academic assignment)
├── hta_method1.png                 # HTA diagram: Photo-first workflow
├── hta_method2.png                 # HTA diagram: Email-first workflow
├── create_presentation.py          # Phase 3 PPTX automation
├── create_presentation_phase5.py   # Phase 5 presentation automation
├── hta_diagrams.py                 # Graphviz HTA diagram generator
├── check_ai_score.py               # AI detection scoring
└── scripts/
    ├── generate_phase4.py          # Phase 4 report automation
    ├── generate_pptx.py            # Phase 4 PPTX builder
    └── create_presentation.py      # Phase 5B presentation automation
```

## Technologies

- **Design**: Figma (lo-fi and hi-fi prototypes)
- **Prototype**: HTML/CSS/JavaScript (interactive browser-based prototype)
- **Automation**: Python (python-pptx, graphviz, python-docx, matplotlib, PyMuPDF)
- **Research**: Structured surveys, contextual inquiry, heuristic evaluation

## Academic Context

IE 48L Special Topics in Human-Computer Interaction Design, Bogazici University, Spring 2026.
