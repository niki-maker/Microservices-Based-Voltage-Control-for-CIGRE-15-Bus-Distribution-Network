# ⚡ Microservices Based Voltage Control for CIGRE-15 Bus Distribution Network

### Volt–VAR Control | Network Reconfiguration  
### CIGRE 15-Bus Distribution System  

---

## 📌 System Overview

This repository presents a Dockerized Active Distribution Management framework developed on the CIGRE 15-Bus Medium Voltage (MV) distribution system.

This implementation focuses exclusively on:

- Volt–VAR Control  
- Network Reconfiguration  

Unlike the IEEE-33 integrated automation framework, this version does **not** include:

- Auto-Reclosure  
- Service Restoration  
- Protection-triggered logic  

The objective is to isolate and evaluate steady-state voltage regulation and topology optimization performance under controlled operating conditions.

The framework is designed for:

- Voltage compliance enforcement (0.95–1.05 p.u.)  
- Switching minimization  
- Radiality preservation  
- Optimization-driven decision making  
- Reproducible containerized deployment  

---

# 🔌 Test System: CIGRE 15-Bus MV Model

### Network Characteristics

- 15 buses  
- Medium-voltage benchmark distribution network  
- Radial structure with normally-open tie switches  
- Distributed load representation  
- Suitable for DER and voltage regulation studies  

This model is widely used in academic research for validating distribution-level optimization strategies.

---

# 🏗 System Architecture

The framework follows a two-layer control structure:

## 1️⃣ Optimization Layer (EnVVarco Core)

Responsible for:

- Continuous bus voltage monitoring  
- Constraint validation (0.95–1.05 p.u.)  
- Multi-objective optimization  
- Switching candidate evaluation  
- Radial topology enforcement  

## 2️⃣ Execution Layer

Responsible for:

- Sequential switching implementation  
- Stabilization delay enforcement  
- Priority-based device activation  
- Power flow re-validation after each switching action  

---

# ⚙️ Core Module

## 🔵 Volt–VAR Control + Network Reconfiguration (EnVVarco)

### Objective

Maintain all bus voltages within permissible limits while minimizing switching operations and preserving radial topology.

---

### Key Functionalities

- Real-time voltage extraction from OpenDSS  
- Undervoltage and overvoltage detection  
- Capacitor / reactor switching decisions  
- Tie-switch-based topology reconfiguration  
- Radiality constraint validation  
- Current threshold validation  
- Multi-objective optimization  

---

### Optimization Strategy

The controller solves a multi-objective problem:

1. Maximize number of buses within 0.95–1.05 p.u.  
2. Minimize number of switching operations  
3. Preserve radial structure  
4. Avoid line current constraint violations  

The selected solution is validated using OpenDSS before execution.

---

# 🐳 Dockerized Microservices Architecture

The framework is deployed using Docker for modularity and reproducibility.

| Container | Responsibility |
|-----------|---------------|
| EnVVarco | Optimization engine |
| Powerflow Service | OpenDSS-based validation |
| Monitoring Layer (Optional) | Visualization support |

---

### Architectural Advantages

- Modular experimentation  
- Easy network model substitution  
- Containerized reproducibility  
- Clear separation between optimization and validation  
- Scalable for future DER integration  

---

# 🔄 Control Flow

```
Normal Operation
↓
Voltage Monitoring
↓
Constraint Violation Detection
↓
Multi-Objective Optimization
↓
Switching Decision Selection
↓
Sequential Execution
↓
Power Flow Validation
↓
Return to Monitoring
```

# 📊 Validation Scenarios

## ✔ Case 1 – Undervoltage Scenario
- Load escalation at remote buses  
- Capacitor activation  
- Optional tie-switch reconfiguration  
- Voltage restored above 0.95 p.u.  

## ✔ Case 2 – Overvoltage Scenario
- Light loading or reactive injection case  
- Reactor switching  
- Topology adjustment if required  
- Voltage restored below 1.05 p.u.  

---

# 🧪 Technologies Used

- Python  
- Flask  
- OpenDSS  
- NumPy  
- Pandas  
- Docker  
- JSON-based inter-service communication  

---

# 🚀 How to Run

```bash
# Clone repository
git clone <your-repository-url>

# Navigate to project directory
cd <repository-name>

# Build containers
docker-compose build

# Start services
docker-compose up
```

# 📌 Key Contributions

- Application of multi-objective Volt–VAR optimization on the CIGRE 15-bus MV system  
- Radiality-aware topology reconfiguration  
- Containerized steady-state distribution automation framework  
- Research-ready modular architecture  
- Scalable foundation for DER-integrated voltage control  

---

# 📌 Conclusion

This repository delivers a modular and research-oriented Active Distribution Management framework applied to the **CIGRE 15-Bus MV distribution system**.

It demonstrates:

- Coordinated Volt–VAR control  
- Radiality-constrained network reconfiguration  
- Dockerized reproducibility  
- Scalable optimization architecture  

It provides a focused steady-state voltage optimization platform suitable for advanced distribution automation research.
