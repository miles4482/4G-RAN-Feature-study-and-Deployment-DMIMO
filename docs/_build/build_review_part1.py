#!/usr/bin/env python3
"""Build D-MIMO Feature Review & Summary Excel (document style, original order)."""

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from xlsx_style import DocBook
from build_review_rest import (
    ch4_config,
    ch4_mml,
    ch4_verify,
    ch5,
    ch6_principles,
    ch6_requirements,
    ch6_config,
    ch6_mml,
    ch6_verify,
    ch7_10,
)

FIG = "/workspace/docs/figures"
OUT = "/workspace/docs/D-MIMO_TDD_eRAN22.1_Feature_Review_and_Summary.xlsx"


def build():
    book = DocBook(
        title="D-MIMO (TDD) Feature Review and Summary",
        subtitle="Document-style restatement of Huawei eRAN D-MIMO (TDD) Feature Parameter Description",
        doc_code="TDLEOFD-111505 / 121501 / 130501",
    )
    cover(book)
    contents(book)
    review_findings(book)
    ch1(book)
    ch2(book)
    ch3(book)
    ch4_principles(book)
    ch4_requirements(book)
    ch4_config(book)
    ch4_mml(book)
    ch4_verify(book)
    ch5(book)
    ch6_principles(book)
    ch6_requirements(book)
    ch6_config(book)
    ch6_mml(book)
    ch6_verify(book)
    ch7_10(book)
    book.save(OUT)
    print("Wrote", OUT)


def cover(book):
    s = book.sheet("Cover", "D-MIMO (TDD) Feature Review & Summary", "Issue 01 (2026-03-10)  |  eRAN22.1 01")
    s.spacer(8)
    s.para("HUAWEI TECHNOLOGIES CO., LTD.", bold=True, color="C7000B")
    s.para("eRAN  |  D-MIMO (TDD) Feature Parameter Description", bold=True)
    s.para("Converted to Excel document style for field review, summary, and sequential deployment planning.")
    s.h2("Document identity")
    s.table([
        ["Field", "Value"],
        ["Source document", "D-MIMO (TDD) Feature Parameter Description"],
        ["Issue / Date", "01  /  2026-03-10"],
        ["Software baseline", "eRAN22.1 01"],
        ["Applicable RAT", "TDD only"],
        ["Pages in source PDF", "95"],
        ["Vendor", "Huawei Technologies Co., Ltd."],
        ["Purpose of this Excel", "Review + summarize every chapter in original order, embed exact figures, and prepare deployment."],
    ])
    s.h2("Features covered (license is on the parent feature)")
    s.table([
        ["Feature ID", "Feature Name", "Type", "Chapter"],
        ["TDLEOFD-111505", "DL D-MIMO", "Commercial parent", "4 Intra-Cell D-MIMO"],
        ["TDLEOFD-11150501", "Distributed MIMO", "Subfeature — basic", "4 Intra-Cell D-MIMO"],
        ["TDLEOFD-11150502", "Intra-BBU D-MIMO", "Subfeature — basic", "4 Intra-Cell D-MIMO"],
        ["TDLEOFD-11150504", "D-MIMO by Macro RRUs", "Subfeature — outdoor macro-macro", "4 Intra-Cell D-MIMO"],
        ["TDLEOFD-11150505", "D-MIMO Within a LampSite Cell", "Subfeature — indoor LampSite", "4 Intra-Cell D-MIMO"],
        ["TDLEOFD-121501", "Inter-eNodeB DL D-MIMO (Trial)", "TRIAL — MoU required", "5 Inter-eNodeB D-MIMO"],
        ["TDLEOFD-130501", "Inter-Cell DL D-MIMO", "Commercial parent", "6 Inter-Cell D-MIMO"],
        ["TDLEOFD-13050101", "Distributed MIMO", "Subfeature — basic", "6 Inter-Cell D-MIMO"],
        ["TDLEOFD-13050102", "Intra-BBU D-MIMO", "Subfeature — basic", "6 Inter-Cell D-MIMO"],
        ["TDLEOFD-13050103", "D-MIMO by Macro RRUs", "Subfeature — outdoor macro-macro", "6 Inter-Cell D-MIMO"],
        ["TDLEOFD-13050104", "D-MIMO by pRRUs", "Subfeature — indoor LampSite", "6 Inter-Cell D-MIMO"],
    ])
    s.note(
        "This source document only provides guidance for feature activation. Feature deployment and gains depend on the live network scenario. To achieve optimal gains, contact Huawei professional service engineers. Cluster IDs must be planned with the Huawei D-MIMO cluster planning tool.",
        "NOTICE",
    )
    s.note(
        "TDLEOFD-121501 Inter-eNodeB DL D-MIMO is a TRIAL feature. It is not for sale in the current version. An MoU with Huawei is required before official application. Huawei is not liable for trial-feature malfunctions. If later converted to commercial, a license fee is required or the trial feature is invalidated on upgrade.",
        "TRIAL",
    )
    s.footer_end()


def contents(book):
    s = book.sheet("00 Contents", "Contents", "Original document sequence is preserved in this workbook.")
    s.para("Sheets follow the source PDF table of contents. Read top to bottom, left to right. Deployment MML and verification sit in the Operation and Maintenance sheets for each feature, matching chapters 4.4, 5.4 and 6.4.")
    s.table([
        ["Sheet", "Source chapter / section", "What you get"],
        ["Cover", "Title page", "Document identity and feature list"],
        ["00 Contents", "Contents", "This map"],
        ["00 Review Findings", "Cross-document review", "Go / no-go summary before you deploy"],
        ["01 Change History", "1 Change History", "eRAN22.1 01 and Draft A deltas"],
        ["02 About This Document", "2 About This Document", "Purpose, RAT, feature/subfeature mapping"],
        ["03 General Principles", "3.1–3.3", "Architecture, calibration, coherent JT + exact figures"],
        ["04-1 Intra-Cell Principles", "4.1–4.2", "Networking, benefits, impacts + Figure 4-1"],
        ["04-2 Intra-Cell Requirements", "4.3", "License, functions, hardware, networking gates"],
        ["04-3 Intra-Cell Data Prep", "4.4.1.1", "All activation / optimization parameters"],
        ["04-4 Intra-Cell MML", "4.4.1.2–4.4.1.3", "Activation and deactivation command examples"],
        ["04-5 Intra-Cell Verify", "4.4.2–4.4.4", "Observation, counters, MAE alarm GUI, faults"],
        ["05 Inter-eNodeB D-MIMO", "5 (all)", "Trial feature: clock, license, extra MML, faults"],
        ["06-1 Inter-Cell Principles", "6.1–6.2", "Intra/Inter-BBP, PCI mod 3 + exact figures"],
        ["06-2 Inter-Cell Requirements", "6.3", "License, exclusive functions, LCOP, RF, cluster rules"],
        ["06-3 Inter-Cell Data Prep", "6.4.1.1", "All activation / optimization parameters"],
        ["06-4 Inter-Cell MML", "6.4.1.2–6.4.1.3", "3-cell activation and deactivation examples"],
        ["06-5 Inter-Cell Verify", "6.4.2–6.4.4", "Verification, monitoring, troubleshooting, alarms"],
        ["07-10 Params Counters Refs", "7–10", "Parameter/counter FAQs, glossary pointer, references"],
    ])
    s.note("Companion workbook: D-MIMO_TDD_eRAN22.1_Step_by_Step_Deployment_Guide.xlsx restates the same content as a field playbook (Phase 0 through rollback).", "TIP")
    s.footer_end()


def review_findings(book):
    s = book.sheet("00 Review Findings", "Review Findings (read this before you deploy)", "Independent review of the source feature document for field deployment.")
    s.h2("1. What D-MIMO is")
    s.para(
        "Distributed MIMO (D-MIMO) is a centralized-scheduling + distributed-RRU feature. Multiple physical antennas transmit in a coordinated way so overlapping coverage becomes useful spatial resolution instead of interference. After a cluster is configured and the feature switch is on, the eNodeB periodically runs (1) inter-RRU channel calibration and (2) coherent joint transmission (SU-JT or MU-JT)."
    )
    s.h2("2. Three networking modes — pick exactly one cluster type per cluster")
    s.table([
        ["Mode", "Feature ID", "PCI rule", "BBU rule", "Introduced", "Switch option"],
        ["Intra-cell D-MIMO", "TDLEOFD-111505", "Same PCI (SFN / pRRU aggregation logical cell)", "Same BBU", "eRAN TDD 11.1", "DmimoJTSwitch"],
        ["Inter-eNodeB D-MIMO (TRIAL)", "TDLEOFD-121501", "Same as intra-cell (intra-cell D-MIMO across BBUs)", "Different BBUs on same Cloud BB; full frequency sync", "eRAN TDD 12.1", "Requires downlink D-MIMO + clock"],
        ["Inter-cell D-MIMO", "TDLEOFD-130501", "Different PCIs, but PCI mod 3 MUST be aligned", "Same BBU (intra-BBP or inter-BBP)", "eRAN TDD 13.1", "InterCellDmimoJTSwitch"],
    ])
    s.h2("3. Deployment verdict")
    s.table([
        ["Topic", "Verdict", "Why it matters on a live network"],
        ["Commercial vs trial", "Deploy 111505 and/or 130501 commercially. Do not deploy 121501 without Huawei MoU.", "Trial features can be invalidated on upgrade and have no commercial warranty."],
        ["Must-have before MML", "License + SFN (intra-cell) or planned PCI/cluster (inter-cell) + compatible BBP/RRU + co-processing board + exclusive functions OFF.", "Wrong order is the main cause of 'Incorrect configurations or unavailable licenses' cluster status."],
        ["Cluster planning", "Mandatory Huawei D-MIMO cluster planning tool. Do not invent cluster IDs.", "Cluster ID uniqueness, mod 6 (multi-carrier), mod 24 (neighbors), even IDs for >4 pRRUs."],
        ["Channel calibration", "Success is the difference between array gain and 'just power overlay'.", "Intra-cell: cal fault → auto revert to SFN. Inter-cell: no revert → only non-coherent JT."],
        ["Energy saving", "Treat RF-channel shutdown, carrier shutdown, low-power, pRRU dormancy, SOC low-power as cluster killers.", "They fail inter-RRU calibration and disable the cluster."],
        ["Best scenario", "Short ISD, LOS between RRUs, UE speed < 5 km/h, heavy overlap.", "Intra-cell: ISD ≤ 150 m outdoor / 30 m LampSite. Inter-cell: ISD ≤ 300 m outdoor; >20% overlap in 6 dB, PRB>10%, BF UEs>20%."],
        ["Do not deploy", "High-speed mobility, extended CP, Massive MIMO introduction, WTTx turbo BF, RRU combine/cascade/ring/load-share, NR-capable RF, BBU3910A, UBRlb, RRU3279 split 4T4R.", "Hard exclusive or unsupported hardware."],
        ["Gains expectation", "Average DL user rate and CEU rate up; RAR success slightly down; access delay slightly up.", "Document does not quote a percentage gain. Gains are scenario-specific."],
    ])
    s.h2("4. Recommended commercial parameter set (from the source, not invented)")
    s.table([
        ["Parameter", "Intra-cell value", "Inter-cell value", "When"],
        ["DMIMOCluster.BfWeightNormalizeMode", "NEBF", "NEBF", "Commercial (edge-optimized MU-JT)"],
        ["CellAlgoSwitch.BfAlgoSwitch", "BfSwitch ON", "BfSwitch ON", "Always"],
        ["CellAlgoSwitch.MuBfAlgoSwitch", "MuBfSwitch ON", "MuBfSwitch ON", "Always"],
        ["CellAlgoSwitch.DMIMOAlgoSwitch", "DmimoJTSwitch ON", "InterCellDmimoJTSwitch ON", "Activation"],
        ["CellAlgoSwitch.EnhChnCalSwitch", "QUICK_CHN_CAL_SWITCH ON", "QUICK_CHN_CAL_SWITCH ON", "Recommended"],
        ["CellAlgoSwitch.CoordinationAlgoSwitch", "MULTI_UE_COORDINATION_OPT_SW ON", "n/a (inter-cell uses DlSchExtSwitch)", "Intra-cell heavy load"],
        ["CellAlgoSwitch.DlSchExtSwitch", "—", "DL_COORD_SCH_EXP_OPT_SW ON", "Inter-cell, low CEU rate"],
        ["CellAlgoSwitch.AvoidInterfSwitch", "—", "AvoidCrsConflictInterfSw ON", "After PCI mod 3 aligned"],
        ["CellPcAlgo.DMSrsPcSinrOffset", "5", "5", "ON; 0 when OFF"],
        ["CellBf.WaitPairingLayerThd", "30 if heavy load + high multi-layer pairing; else 0", "—", "Intra-cell TMA"],
        ["CellPdcchAlgo.PdcchBfGainOffset", "0 if STRATEGYBASEDONCOVERAGE", "—", "Intra-cell"],
        ["CellPdcchAlgo.PDCCHAggLvlAdaptStrage", "—", "STRATEGYBASEDONCAPACITY", "Inter-cell"],
        ["CellBfMimoParaCfg.TmAccelerationSwitch", "INITIAL_ACCESS_TO_BF", "—", "Intra-cell"],
        ["CellDlschAlgo.DlHighLoadSdmaThdOffset", "10 if D-MIMO+quick BF; 5 if D-MIMO only (micro-micro)", "—", "Not used in macro-macro"],
        ["DMIMOAlgo.DmimoA3RsrpOffset", "—", "-12 (actual –6 dB)", "Inter-cell"],
        ["DMIMOAlgo.DmimoJtCellRsrpDiffThld", "—", "-20 (actual –10 dB)", "Inter-cell"],
        ["SRSCfg.RsvResSrsPeriod1 / 2", "—", "10 ms / 20 ms", "Inter-cell; same in cluster"],
        ["NCellSrsMeasPara", "—", "SrsAutoNCellMeasSwitch=ON, A3 offset=-12, NCellSrsTimeMeasSwitch ON", "Inter-cell"],
        ["CellCqiAdjAlgo.CellCqiAdjSchCntThld", "—", "50", "Inter-cell"],
        ["CoProcRes.BundlingClusterType", "DMIMO (or ADAPTIVE if mixed coordination)", "DMIMO (or ADAPTIVE)", "WorkMode dedicated LCOP for inter-cell"],
    ])
    s.h2("5. License models")
    s.table([
        ["Feature ID", "Name", "Model", "Sales unit"],
        ["TDLEOFD-111505", "DL D-MIMO", "LT1SDLMIMO00", "per Cell"],
        ["TDLEOFD-121501", "Inter-eNodeB DL D-MIMO (Trial)", "LT1STDLDMC00", "per eNodeB"],
        ["TDLEOFD-130501", "Inter-Cell DL D-MIMO", "LT1SICDLDM00", "per Cell"],
    ])
    s.note("License insufficiency can block cell activation. Check the 'Is Cell Activation Affected by License Insufficiency' column in License Control Item Lists.", "NOTICE")
    s.h2("6. Sequence you must follow (matches source O&M chapters)")
    s.bullets([
        "Confirm RAT is TDD and the chosen feature is licensed (parent feature).",
        "Confirm hardware, RF, cell bandwidth/band, and networking restraints.",
        "Turn OFF mutually exclusive functions; understand function impacts (especially energy saving and DRX).",
        "Activate prerequisite functions (SFN + BF + MU-BF for intra-cell; BF + MU-BF for inter-cell; downlink D-MIMO for inter-eNodeB).",
        "Plan cluster IDs (and PCI mod 3 for inter-cell) with Huawei tools.",
        "Configure CoProcRes → DMIMOCluster → DMIMOClusterCell → algorithm switches → recommended offsets.",
        "Activate with MML or MAE-Deployment. Observe DSP CELL / DSP DMIMOCLUSTERCELL / DSP DMIMOCALIBRATION (and DSP CELLCALIBRATION for inter-cell).",
        "Monitor pairing counters and CEU throughput. Configure MAE threshold alarm on L.CellSectorEqpt.UNA.Dur.Cali.",
        "If rolling back, follow the deactivation MML in reverse of cluster membership, then restore offsets.",
    ], numbered=True)
    s.footer_end()


def ch1(book):
    s = book.sheet("01 Change History", "1  Change History", "Source pages 7. Technical vs editorial changes only.")
    s.para("This chapter describes changes not included in Parameters, Counters, Glossary, and Reference Documents: technical changes (functions and parameters) and editorial changes (documentation).")
    s.h2("1.1  eRAN22.1 01 (2026-03-10)")
    s.para("This issue does not include any changes.")
    s.h2("1.2  eRAN22.1 Draft A (2025-12-31)")
    s.para("This issue introduces the following changes to eRAN21.1 01 (2025-03-10).")
    s.h3("Technical Changes")
    s.para("None")
    s.h3("Editorial Changes")
    s.bullets([
        "Added an impact relationship between inter-cell downlink D-MIMO and downlink FDD+TDD CA. See 6.3.2.3 Function Impacts. (Inter-cell downlink D-MIMO does not work in SCells.)",
        "Added an impact relationship between downlink D-MIMO and precise AMC. See 4.3.2.3 Function Impacts. (Precise AMC does not take effect for D-MIMO UEs in the joint transmission state.)",
    ])
    s.footer_end()


def ch2(book):
    s = book.sheet("02 About This Document", "2  About This Document", "Source pages 8–10.")
    s.h2("2.1  General Statements")
    s.h3("Purpose")
    s.para("This document is intended to acquaint readers with:")
    s.bullets([
        "The technical principles of features and their related parameters",
        "The scenarios where these features are used, the benefits they provide, and the impact they have on networks and functions",
        "Requirements that must be met before feature activation",
        "Parameter configuration required for feature activation, verification of feature activation, and monitoring of feature performance",
    ])
    s.note(
        "This document only provides guidance for feature activation. Feature deployment and feature gains depend on the specifics of the network scenario where the feature is deployed. To achieve optimal gains, contact Huawei professional service engineers. Functions mentioned in this document work properly only when enabled in the specified applicable scenarios (such as RAT and networking).",
        "NOTICE",
    )
    s.h3("Software Interfaces")
    s.para("Any parameters, alarms, counters, or managed objects (MOs) described in this document apply only to the corresponding software release. For future software releases, refer to the corresponding updated product documentation.")
    s.h3("Trial Features")
    s.para(
        "Trial features are not yet ready for full commercial release (for example industry-chain compatibility). They can be used for testing or commercial network trials. Anyone who desires to use trial features shall contact Huawei and enter into a memorandum of understanding (MoU) prior to official application. Trial features are not for sale in the current version but customers may try them for free. Customers acknowledge risk due to absence of commercial testing. Huawei is not liable for trial feature malfunctions or losses. Huawei does not promise that problems will be resolved in the current version. Huawei may convert trial features into commercial features in later R/C versions; customers shall then pay a licensing fee. If a customer fails to purchase such a license, the trial feature(s) will be invalidated automatically when the product is upgraded."
    )
    s.h3("Feature Differences Between RATs")
    s.para("Unless otherwise stated, descriptions in this document apply to all RATs covered here. If a description does not apply to all RATs, the specific RAT is stated. Example: “TDD cells are compatible with enhanced MU-MIMO” means the function cannot be used in non-TDD cells.")
    s.h2("2.2  Applicable RAT")
    s.para("This document applies to TDD.", bold=True)
    s.h2("2.3  Features in This Document")
    s.table([
        ["Feature ID", "Feature Name", "Chapter/Section"],
        ["TDLEOFD-111505", "DL D-MIMO", "4 Intra-Cell D-MIMO"],
        ["TDLEOFD-11150501", "Distributed MIMO", "4 Intra-Cell D-MIMO"],
        ["TDLEOFD-11150502", "Intra-BBU D-MIMO", "4 Intra-Cell D-MIMO"],
        ["TDLEOFD-11150504", "D-MIMO by Macro RRUs", "4 Intra-Cell D-MIMO"],
        ["TDLEOFD-11150505", "D-MIMO Within a LampSite Cell", "4 Intra-Cell D-MIMO"],
        ["TDLEOFD-121501", "Inter-eNodeB DL D-MIMO (Trial)", "5 Inter-eNodeB D-MIMO"],
        ["TDLEOFD-130501", "Inter-Cell DL D-MIMO", "6 Inter-Cell D-MIMO"],
        ["TDLEOFD-13050101", "Distributed MIMO", "6 Inter-Cell D-MIMO"],
        ["TDLEOFD-13050102", "Intra-BBU D-MIMO", "6 Inter-Cell D-MIMO"],
        ["TDLEOFD-13050103", "D-MIMO by Macro RRUs", "6 Inter-Cell D-MIMO"],
        ["TDLEOFD-13050104", "D-MIMO by pRRUs", "6 Inter-Cell D-MIMO"],
    ])
    s.para("The following table lists the mapping between subfeatures and parent features. The license is deployed on the parent features.")
    s.table([
        ["Parent Feature", "Subfeature", "Scenario"],
        ["TDLEOFD-111505 DL D-MIMO", "TDLEOFD-11150501 Distributed MIMO", "Basic functions"],
        ["TDLEOFD-111505 DL D-MIMO", "TDLEOFD-11150502 Intra-BBU D-MIMO", "Basic functions"],
        ["TDLEOFD-111505 DL D-MIMO", "TDLEOFD-11150504 D-MIMO by Macro RRUs", "Outdoor macro-macro scenarios"],
        ["TDLEOFD-111505 DL D-MIMO", "TDLEOFD-11150505 D-MIMO Within a LampSite Cell", "Indoor LampSite scenarios"],
        ["TDLEOFD-130501 Inter-Cell DL D-MIMO", "TDLEOFD-13050101 Distributed MIMO", "Basic functions"],
        ["TDLEOFD-130501 Inter-Cell DL D-MIMO", "TDLEOFD-13050102 Intra-BBU D-MIMO", "Basic functions"],
        ["TDLEOFD-130501 Inter-Cell DL D-MIMO", "TDLEOFD-13050103 D-MIMO by Macro RRUs", "Outdoor macro-macro scenarios"],
        ["TDLEOFD-130501 Inter-Cell DL D-MIMO", "TDLEOFD-13050104 D-MIMO by pRRUs", "Indoor LampSite scenarios"],
    ])
    s.footer_end()


def ch3(book):
    s = book.sheet("03 General Principles", "3  General Principles", "Source pages 11–19. Exact figures from the PDF.")
    s.h2("3.1  Overview")
    s.h3("Background")
    s.para(
        "As the number of users and traffic volume increase sharply, site density rises. Densification improves capacity but increases overlapping coverage and degrades cell-edge experience. Distributed MIMO (D-MIMO) is introduced to mitigate interference from neighboring cells in overlapping coverage areas and improve the experience of cell edge UEs (CEUs) and the downlink user experience of the whole cell."
    )
    s.h3("Introduction")
    s.para(
        "D-MIMO is based on a centralized scheduling + distributed RRU architecture (Figure 3-1). It uses multiple physical antennas to transmit data in a coordinated way. Dispersed antenna locations significantly improve spatial channel resolution. This offers high antenna array gains and interference suppression gains while improving downlink user experience and CEU experience."
    )
    s.figure(f"{FIG}/fig_3_1a_without_with_dmimo.png", "Figure 3-1 (part)  Without D-MIMO vs With D-MIMO — interference converted into energy")
    s.figure(f"{FIG}/fig_3_1b_dmimo_architecture.png", "Figure 3-1  D-MIMO — distributed antennas (×10 m to ×100 m) with centralized scheduling")
    s.para(
        "D-MIMO coordinates scheduling on a per D-MIMO cluster basis. A D-MIMO cluster is a set of physical cells (generated by RRUs) that compose one or more cells. D-MIMO requires that D-MIMO cluster data be planned, including the RRUs in a D-MIMO cluster and the cluster ID. For the planning method, contact Huawei engineers."
    )
    s.h3("Category — three networking modes")
    s.table([
        ["Feature ID", "Feature Name", "Networking", "Introduced in", "Chapter"],
        ["TDLEOFD-111505", "DL D-MIMO", "Intra-cell: multiple RRUs on the same BBU form a logical cell; cluster for that logical cell. Cells in the cluster MUST have the same PCI.", "eRAN TDD 11.1", "4 Intra-Cell D-MIMO"],
        ["TDLEOFD-121501", "Inter-eNodeB DL D-MIMO (Trial)", "Intra-cell style: RRUs on different BBUs of the same Cloud BB network form a cluster.", "eRAN TDD 12.1", "5 Inter-eNodeB D-MIMO"],
        ["TDLEOFD-130501", "Inter-Cell DL D-MIMO", "Inter-cell: multiple RRUs on the same BBU form a cluster. Cells in the cluster have DIFFERENT PCIs.", "eRAN TDD 13.1", "6 Inter-Cell D-MIMO"],
    ])
    s.h3("Working process")
    s.figure(f"{FIG}/fig_3_2_dmimo_process.png", "Figure 3-2  D-MIMO process")
    s.bullets([
        "D-MIMO configuration: involves D-MIMO clusters, feature switches, and related parameters.",
        "Inter-RRU channel calibration: distributed RRUs transmit with different phases. After calibration, signals from multiple RRU channels combine in the same phase on the UE side.",
        "Coherent joint transmission: algorithm identifies JT UEs (SU-JT or MU-JT) in the cluster and performs coherent JT, improving downlink and cell-edge experience.",
    ], numbered=True)
    s.para("After a cluster is configured and D-MIMO is enabled, the eNodeB periodically initiates intra-cluster inter-RRU channel calibration and coherent joint transmission. Control options:")
    s.bullets([
        "Intra-cell D-MIMO: DmimoJTSwitch option of CellAlgoSwitch.DMIMOAlgoSwitch",
        "Inter-cell D-MIMO: InterCellDmimoJTSwitch option of CellAlgoSwitch.DMIMOAlgoSwitch",
    ])
    s.h2("3.2  Inter-RRU Channel Calibration")
    s.para(
        "In the distributed RRU architecture, RRUs in different locations transmit with different phases. Inter-RRU channel calibration ensures RF channels of a cluster have the same delay so signals combine in phase on the UE side. The eNodeB adjusts the transmit/receive response ratio of each RF channel to be consistent with the reference RF channel."
    )
    s.figure(f"{FIG}/fig_3_3_inter_rru_calibration.png", "Figure 3-3  Principles of inter-RRU channel calibration (α1 compensation)")
    s.para("Procedure:")
    s.bullets([
        "The intra-RRU channel calibration algorithm performs channel calibration for each RRU.",
        "The inter-RRU channel calibration algorithm is used. RRUs transmit calibration sequences to each other over the air interface and calculate the calibration coefficient.",
        "Channel compensation is performed so transmit/receive response ratios of all RF channels in the cluster are consistent.",
    ], numbered=True)
    s.note("2T2R RRUs and pRRUs do not support intra-RRU channel calibration. Therefore only steps 2 and 3 are required for 2T2R RRUs and pRRUs.", "NOTE")
    s.para("If a channel calibration fault occurs in a cell in the cluster:")
    s.bullets([
        "Intra-cell D-MIMO: D-MIMO cells automatically revert to SFN configurations.",
        "Inter-cell D-MIMO: D-MIMO cells do not revert. UEs can obtain power superimposition gains but cannot obtain array gains.",
    ])
    s.h3("Related optimization (EnhChnCalSwitch)")
    s.para(
        "In eRAN TDD 11.1 or earlier, intra-RRU channel calibration is initiated every 30 minutes. In eRAN TDD 12.0, it is initiated every 10 s if QUICK_CHN_CAL_SWITCH of CellAlgoSwitch.EnhChnCalSwitch is selected; otherwise still every 30 minutes. It is recommended that this option be selected in D-MIMO scenarios. The setting does not take effect on 2T2R RRUs and pRRUs."
    )
    s.h2("3.3  Coherent Joint Transmission")
    s.para("Multiple RRUs jointly transmit to the same UE; signals are superimposed in phase on the UE side, increasing downlink received strength.")
    s.figure(f"{FIG}/fig_3_4_coherent_jt.png", "Figure 3-4  Coherent joint transmission — centralized scheduling in a D-MIMO cluster")
    s.para("Works as follows:")
    s.bullets([
        "Identifying JT UEs: SRS is used to obtain uplink RSRP at each RRU and calculate downlink equivalent RSRP in intra-cell scenarios. If the working RRU list for a UE contains two or more RRUs, the UE is a JT UE. In inter-cell scenarios, SRS and event A3 are used; if both are available, event A3 RSRP is preferred.",
        "Calculating beamforming weighting values from uplink channel coefficients (SRS). Enable target SRS SINR optimization for BF UEs via CellPcAlgo.DMSrsPcSinrOffset (>0 adds to SrsPcSinrTarget). This may slightly increase RRC Connection Reconfiguration messages and the service drop rate.",
        "Performing SU-JT or MU-JT in each TTI.",
    ], numbered=True)
    s.figure(f"{FIG}/fig_3_5_jt_process.png", "Figure 3-5  JT process (SU-JT then MU-JT)")
    s.figure(f"{FIG}/fig_3_6_su_jt.png", "Figure 3-6  SU-JT — multiple RRUs transmit coherently to one UE")
    s.figure(f"{FIG}/fig_3_7_mu_jt.png", "Figure 3-7  MU-JT — orthogonal weights, resource multiplexing")
    s.para("Set DMIMOCluster.BfWeightNormalizeMode based on site conditions. NEBF = optimal for paired BF UEs at the D-MIMO cell edge. PEBF = optimal at center / medium distance. Set NEBF in commercial use.")
    s.figure(f"{FIG}/fig_3_8_mu_jt_example.png", "Figure 3-8  MU-JT example — V0 paired with U0 on RBG 0–3 and with U1 on RBG 4–7")
    s.para(
        "In Figure 3-8, U0 and U1 are already scheduled; V0 is to be paired. Before MU-JT only U0 and U1 use RBG 0–7. After pairing, three UEs share the same frequency resources, improving spectral efficiency."
    )
    s.para(
        "When cell-level TMA is enabled in non-inter-cell D-MIMO scenarios, more UEs in heavy-load cells enter BF for MU pairing. Controlled by CellBf.WaitPairingLayerThd (0 = off; >0 and DL PRB usage above internal threshold AND BF UEs to schedule per TTI ≥ WaitPairingLayerThd × 0.1)."
    )
    s.para(
        "Advanced Multi-User Coordination further increases MU-JT pairing rate and DL capacity in intra-cell D-MIMO under heavy load. Switch: MULTI_UE_COORDINATION_OPT_SW of CellAlgoSwitch.CoordinationAlgoSwitch. See Beamforming (TDD)."
    )
    s.footer_end()


def ch4_principles(book):
    s = book.sheet("04-1 Intra-Cell Principles", "4  Intra-Cell D-MIMO  —  Principles and Network Analysis", "Source pages 20–23. Feature TDLEOFD-111505.")
    s.h2("4.1  Principles")
    s.para(
        "Intra-cell D-MIMO is developed based on SFN. Multiple RRUs connected to a single BBU with overlapping coverage, and their corresponding physical cells, constitute a logical cell. A D-MIMO cluster is set up for that logical cell. The cluster is equivalent to an SFN cell, so the number of RRUs in the cluster equals the number of RRUs serving the SFN cell. Cells in a D-MIMO cluster have the same PCI. RRUs connect to the same BBU."
    )
    s.figure(f"{FIG}/fig_4_1_intracell_networking.png", "Figure 4-1  Networking diagram of intra-cell D-MIMO (source figure on p.21)")
    s.para("Required SFN scenarios:")
    s.bullets(["Outdoor macro-macro SFN scenarios", "Indoor LampSite SFN scenarios"])
    s.para("Typical areas:")
    s.bullets([
        "Outdoor hotspot areas: densely populated urban areas, CBDs, campus",
        "Indoor densely populated areas: stadiums, airports, railway stations, dining halls, auditoriums",
    ])
    s.note("In this document, “macro-macro networking” refers to a combination of macro RRUs. See RF Modules.", "NOTE")
    s.h2("4.2.1  Benefits")
    s.h3("Most beneficial scenarios — all of the following")
    s.bullets([
        "Inter-site distance does not exceed 150 m outdoors or 30 m in LampSite, with significant coverage overlap.",
        "Interference between neighboring cells is severe in LampSite (large open intra-frequency areas such as transport hubs and exhibition centers).",
        "Line-of-sight (LOS) propagation between RRUs.",
        "UE speed lower than 5 km/h.",
    ])
    s.h3("Benefits")
    s.bullets([
        "Outdoor macro-macro SFN: increases average DL user-experienced rate and DL CEU-experienced rate.",
        "Outdoor micro-micro SFN: same.",
        "Indoor LampSite SFN: increases average DL UE-experienced rate.",
        "Combined with advanced multi-user coordination: more spatial multiplexing; increases cell DL traffic volume and spectral efficiency when load is heavy.",
    ])
    s.h2("4.2.2  Impacts")
    s.bullets([
        "Slightly decreases RAR success rate but increases UE access delay.",
        "With adaptive TM switching in SFN cells, more UEs switch TM3 → TM8, so a greater proportion use BF (including dual-stream). Dual-stream TM3 reports per-stream CQI; dual-stream TM8 reports CQI for one stream → average reported CQI increases.",
        "Dual-stream BF UEs must switch to single-stream before pairing → rank-1 BF PRB counter rises; rank-2 BF PRB proportion may fall if many dual-stream UEs pair.",
        "More UEs scheduled per TTI when many UEs wait, but needs enough CCEs. If CCE sufficient: DL CCE use up, UL CCE availability down, UL CCE fail probability up. If CCE insufficient: more CCE allocation failures per TTI.",
        "Joint-scheduling UEs receive D-MIMO gains; pairing rate improves; BLER may slightly fluctuate with MCS changes.",
        "SFN/D-MIMO: CRS jointly transmitted, PDSCH independently transmitted for independent-scheduling UEs → possible RI mismatch. 2T/4T: SfnDlRblerOptSwitch forces rank 1 (~1 s to identify). 8T: TMA to TM8 (~2 s). Short mismatch windows can drop MCS/traffic.",
    ])
    s.footer_end()


def ch4_requirements(book):
    s = book.sheet("04-2 Intra-Cell Requirements", "4.3  Intra-Cell Requirements", "Source pages 23–33. Complete these gates before any MML.")
    s.h2("4.3.1  Licenses")
    s.table([
        ["Feature ID", "Feature Name", "Model", "Sales Unit"],
        ["TDLEOFD-111505", "DL D-MIMO", "LT1SDLMIMO00", "per Cell"],
    ])
    s.note("Insufficiency of certain feature licenses affects cell activation. See License Control Item Lists.", "NOTICE")
    s.h2("4.3.2  Functions — rule")
    s.para("Before activating, ensure prerequisite functions are activated and mutually exclusive functions are deactivated. Understand function impacts. Deactivate impacted functions using the relevant feature documents.")
    s.h3("4.3.2.1  Prerequisite functions")
    s.table([
        ["Function", "Switch", "Reference", "Description"],
        ["Single-stream beamforming", "BfSwitch of CellAlgoSwitch.BfAlgoSwitch", "Beamforming (TDD)", "Outdoor macro-macro: prerequisite. LampSite: license not required, but BfSwitch MUST still be selected."],
        ["MU beamforming", "MuBfSwitch of CellAlgoSwitch.MuBfAlgoSwitch", "Beamforming (TDD)", "Outdoor macro-macro: prerequisite. LampSite: license not required, but MuBfSwitch MUST still be selected."],
        ["SFN", "CellAlgoSwitch.SfnUlSchSwitch and SfnDlSchSwitch", "SFN", "N/A (required architecture)"],
    ])
    s.h3("4.3.2.2  Mutually exclusive functions — must be OFF")
    s.table([
        ["Function", "Switch", "Reference", "Note"],
        ["Extended CP", "Cell.UlCyclicPrefix / Cell.DlCyclicPrefix", "Extended CP", "None"],
        ["High speed mobility", "Cell.HighSpeedFlag", "High Speed Mobility", "None"],
        ["Out-of-band relay", "CellAlgoSwitch.RelaySwitch", "Relay", "None"],
        ["Massive MIMO introduction", "SrvBasedSRSAdjAlgo of SRSCfg.SrsCfgPolicySwitch (others free of switch)", "Massive MIMO Basics (TDD)", "None"],
        ["Uplink SU-MIMO", "CellAlgoSwitch.UlSuMimoAlgoSwitch", "MIMO", "None"],
        ["WTTx turbo beamforming", "TurboBfSwitch of MuBfAlgoSwitch; AntSelEnhanceBfSwitch of BfAlgoSwitch", "WTTx Turbo Beamforming (TDD)", "None"],
        ["SRS interference avoidance", "WTTxSRSIntrfAvoidanceSw of SRSCfg.SrsCfgPolicySwitch", "Massive MIMO Optimization in WTTx Scenarios (TDD)", "None"],
        ["Energy saving based on proactive scheduling", "SymbolPwrSaving.TrigBndlSchDlAvgPrbThld", "Energy Conservation and Emission Reduction", "Cannot be enabled together with intra-cell D-MIMO"],
        ["RF channel dynamic muting", "RF_CHN_DYN_MUTING_SW of CellRfChnDynMuting.RfChnDynMutingAlgoSwitch", "Energy Conservation…", "None"],
        ["Dynamic TM3R2↔TM9R4 threshold adapt", "TM3R2_TO_TM9R4_THLD_ADAPT_SW of CellBfMimoParaCfg.BfMimoAlgoOptSwitch", "Beamforming (TDD)", "Does not take effect if Cell.MultiRruCellFlag = BOOLEAN_TRUE"],
        ["Dynamic TM3R2↔dual-stream BF threshold adapt", "TM3R2_TO_DUALBF_THLD_ADAPT_SW of CellBfMimoParaCfg.BfMimoAlgoOptSwitch", "Beamforming (TDD)", "Does not take effect if Cell.MultiRruCellFlag = BOOLEAN_TRUE"],
        ["Intelligent Multi-Beam of 8T8R", "None", "Intelligent Multi-Beam of 8T8R (TDD)", "None"],
        ["Adaptive MU beamforming", "CellBf.AdaptMubfStartThld ≠ 255", "Beamforming (TDD)", "None"],
        ["MuteUpptsSymbForSrsSw", "MuteUpptsSymbForSrsSw of SRSCfg.SrsCfgPolicySwitch", "None", "Cannot be on together with intra-cell D-MIMO"],
        ["Single-cell intelligent PC energy saving", "INTEL_INTRA_CELL_PC_SW of CellPwrSavingAlgo.PwrSavingAlgoSwitch", "Energy Conservation…", "None"],
    ])
    s.h3("4.3.2.3  Function impacts")
    s.table([
        ["Function", "Switch", "Impact"],
        ["RF channel intelligent shutdown", "CellRfShutdown.RfShutdownSwitch", "Inter-RRU channel calibration WILL FAIL if enabled."],
        ["Intelligent power-off of co-coverage carriers", "CellShutdown.CellShutdownSwitch", "If a cell in an inter-cell cluster enters this state, cal fails and the inter-cell cluster is disabled. (Listed in intra-cell impacts chapter.)"],
        ["Low power consumption mode", "CellLowPower.LowPwrSwitch", "Cell cannot serve as cooperating cell; inter-RRU cal may fail."],
        ["MU beamforming", "MuBfSwitch", "When D-MIMO is on, MU spatial multiplexing is done by D-MIMO; MU-BF counters/monitoring are no longer measured."],
        ["DRX", "CellDrxPara.DrxAlgSwitch", "UEs do not send SRS in DRX sleep → BF weights not updated promptly → D-MIMO performance deteriorates."],
        ["Dynamic Power Sharing Between LTE Carriers", "LTE_DYN_POWER_SHARING_SW", "D-MIMO serves one cell from multiple RF modules; power sharing is per physical cell / one RF module. Combined, DPS cannot provide maximum gains."],
        ["SFN", "—", "Interference suppression uses D-MIMO; SFN inter-RRU coordinated BF counters/monitoring no longer measured."],
        ["Adaptive switching BF ↔ MIMO", "CellBfMimoParaCfg.BfMimoAdaptiveSwitch", "Light load: TMA for optimal single-link. Heavy load: do not use that; use cell-level TMA when multi-layer pairing is available."],
        ["Preferential RBG alloc type 1", "CellDlschAlgo.RbgAllocStrategy = TYPE1_FIRST", "Does not take effect when D-MIMO is enabled."],
        ["MCS selection for small-data", "SmallPktMcsSelectAlgoSw of CellAlgoSwitch.DlSchSwitch", "Does not take effect for D-MIMO UEs in JT state."],
        ["pRRU deep dormancy (intelligent)", "EnodebMpruEs.MpruDormancyDlEarfcn", "Downlink D-MIMO no longer takes effect when some pRRUs enter dormancy."],
        ["Precise AMC", "PreciseAmcSwitch of CellAlgoSwitch.EmimoSwitch", "Does not take effect for D-MIMO UEs in JT state. (Added in eRAN22.1 Draft A.)"],
    ])
    s.h2("4.3.3  Hardware")
    s.para("See 3900 & 5900 Series Base Station Product Documentation for compatibility.")
    s.h3("Base station models")
    s.bullets([
        "3900 and 5900 series (macro)",
        "DBS3900 LampSite and DBS5900 LampSite",
        "TDLEOFD-11150501 Distributed MIMO: macro and LampSite",
        "TDLEOFD-11150502 Intra-BBU D-MIMO: macro",
        "TDLEOFD-11150504 D-MIMO by Macro RRUs: macro",
        "TDLEOFD-11150505 D-MIMO Within a LampSite Cell: LampSite",
    ])
    s.h3("Table 4-1  Cell specifications supported by BBPs")
    s.table([
        ["Board", "Tx/Rx mode", "Maximum number of 10 MHz or 20 MHz cells"],
        ["UBBPd6 or UBBPd9", "2T2R or 4T4R", "6 if not co-processing board; 3 if also co-processing board"],
        ["UBBPd6 or UBBPd9", "8T8R", "6 if not co-processing; 3 if also co-processing"],
        ["UBBPe4", "2T2R or 4T4R", "6. Cannot function as co-processing board."],
        ["UBBPe4", "8T8R", "3. Cannot function as co-processing board."],
        ["UBBPe6", "2T2R or 4T4R", "12"],
        ["UBBPe6", "8T8R", "6"],
    ])
    s.note("UBRlb does not support D-MIMO. Outdoor BBU3910A does not support D-MIMO or enhanced channel calibration. Cells in a cluster can be on different BBPs but MUST be in the same BBU. UBBPe4/d6/d9/e6 series required to set up cells and connect RRUs.", "NOTICE")
    s.h3("Co-processing board")
    s.bullets([
        "Required to calculate D-MIMO co-processing resources. Use UBBPd6, UBBPe6, or UBBPd9.",
        "UBBPd: max 72 antennas and 6 clusters (e.g. 8T 3 clusters × 3 RRU; 4T 6×3; 2T 6×6).",
        "UBBPe: max 144 antennas and 12 clusters (e.g. 8T 6×3; 4T 12×3; 2T 12×6).",
        "4T and 8T clusters can share co-processing resources.",
        "BBU3900: co-processing board MUST be in slot 2 or 3. BBU3910: no slot constraint.",
        "Cannot be a dedicated scheduling board (signaling-only BBP).",
        "UBBPd cannot be co-processing if 4 or more physical cells are bound to it.",
        "UBBPe cannot be co-processing if ≥4 8T8R 20 MHz physical cells without CPRI compression, or ≥7 4T4R/8T8R physical cells in other scenarios.",
        "FDD+TDD BBP cannot be a D-MIMO co-processing board.",
        "If a BBP is co-processing, you cannot set up both 4T4R and 8T8R cells on that BBP.",
        "Co-processing board uses CellFrameOffset.FrameOffset of the corresponding BBP. Cluster frame offset must match the co-processing board for coordinated scheduling.",
    ])
    s.h3("RF modules")
    s.table([
        ["Category", "RRU Model", "Channels", "Band", "Scenario"],
        ["Macro RRU", "RRU3252", "4T4R", "Band 38", "Outdoor macro-macro"],
        ["Macro RRU", "RRU3256", "4T4R", "Band 42 or 43", "Outdoor macro-macro"],
        ["Macro RRU", "RRU3278", "8T8R", "Band 42 or 43", "Outdoor macro-macro"],
        ["Macro RRU", "RRU3279", "8T8R", "Band 38", "Outdoor macro-macro"],
        ["pRRU", "pRRU3901", "2T2R", "2.3 GHz", "Indoor LampSite"],
        ["pRRU", "pRRU3911", "2T2R", "2.3 / 2.5 / 2.6 GHz", "Indoor LampSite"],
    ])
    s.note("RRU3279 serving two 4T4R sectors by splitting does NOT support D-MIMO. NR-capable RF modules do NOT support D-MIMO. RRUs incapable of D-MIMO do not support enhanced channel calibration.", "WARNING")
    s.para("Constraints:")
    s.bullets([
        "RRUs cannot be combined.",
        "Macro RRUs cannot be cascaded.",
        "RRUs cannot use load sharing networking.",
        "RRUs cannot use ring networking.",
        "RRUs cannot be cascaded in macro-macro scenarios.",
        "Remote RRU deployment distance cannot exceed 20 km.",
        "No more than three carriers on the same RRU.",
        "In FDD+TDD RRU cascading, the upper-level RRU can be TDD, not FDD.",
    ])
    s.h3("Cells")
    s.bullets([
        "Bandwidth must be 10 MHz or 20 MHz.",
        "Outdoor: bands 38, 40, 42, 43.",
        "Indoor LampSite: 2.3 / 2.5 / 2.6 GHz.",
    ])
    s.h3("LampSite configuration restraints")
    s.bullets([
        "D-MIMO cannot be used in pRRU combination scenarios.",
        "pRRUs in the same cluster must connect to the same RHUB or two neighboring RHUBs on the same CPRI link.",
        "No extender on the Ethernet cable from RHUB to pRRU.",
    ])
    s.h2("4.3.4  Networking")
    s.para("In inter-BBU scenarios, D-MIMO is supported only in LTE TDD single-mode or LTE TDD/TD-SCDMA dual-mode networking.")
    s.h3("Planning of SFN networking")
    s.bullets([
        "In SFN scenarios, D-MIMO can be deployed directly.",
        "In non-SFN scenarios, networking must be planned using a D-MIMO cluster planning tool.",
    ])
    s.h3("Planning of D-MIMO cluster IDs (Huawei tool — mandatory)")
    s.bullets([
        "Basic: cells from more than one RRU or more than two pRRUs form a cluster.",
        "Each cluster ID must have a unique mod 6 value for each RRU in multi-carrier scenarios.",
        "Cluster ID must be unique within a BBU.",
        "IDs of physically neighboring clusters must not have the same mod 24 value.",
        "Macro-macro: 2 to 4 RRUs per cluster.",
        "LampSite: 3 to 6 pRRUs. If multiple carriers, LTE TDD carrier count AND bandwidth must be the same for all pRRUs in the cluster, whether or not D-MIMO is on for all carriers. If more than four pRRUs, cluster ID must be even.",
    ])
    s.footer_end()


if __name__ == "__main__":
    build()
