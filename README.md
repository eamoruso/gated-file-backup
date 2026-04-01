# Title
Integrity‑Gated File Monitoring for Ransomware Resilience and Real‑Time Recovery
# Abstract
File‑encrypting ransomware continues to pose a significant threat to endpoint data despite advances in signature‑based detection 
and behavior‑driven defenses. Rather than attempting early malware classification or execution prevention, 
this work examines whether inexpensive file integrity validation can serve as an effective control point for 
mitigating data loss under modern ransomware behaviors. We present a Windows file‑monitoring system that observes 
filesystem modification events and applies integrity‑gated selective backup, preserving recoverable file versions 
while suppressing backups of corrupted outputs. The system validates modified files using a combination of header 
“magic number” checks and lightweight, format‑aware structural invariants. To address ransomware that employs partial 
or intermittent encryption while preserving file headers, we incorporate low‑overhead spot‑entropy sampling across 
small, deterministic file regions. Process‑aware allow-list suppresses entropy analysis for trusted applications to 
reduce false positives under common workloads. Backups of validated files are stored in a hardened repository with 
obfuscated filenames, and monitoring is implemented using asynchronous directory notifications to minimize missed 
events under sustained I/O. We evaluate a prototype implementation against in‑the‑wild ransomware samples exhibiting 
both full‑file and partial encryption strategies. The system preserved 100% of files targeted by ransomware performing 
complete file encryption and achieved approximately 90% protection against header‑preserving partial encryption 
patterns. Performance evaluation shows that integrity‑gated backup can operate continuously on Windows endpoints 
with minimal CPU and memory overhead and negligible impact on user workflows. These results demonstrate that 
integrity‑based backup gating can provide measurable ransomware impact mitigation under realistic user‑level threat models, complementing prevention‑oriented defenses without relying on heavyweight behavioral analysis.
# Authors
Edward L. Amoruso - Paul E. Amoruso - Cliff C. Zou 
