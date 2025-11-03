import express from "express";
import cors from "cors";
import bodyParser from "body-parser";
import { Document, Packer, Paragraph, HeadingLevel, TextRun } from "docx";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors());
app.use(bodyParser.json());

app.post("/final-report", async (req, res) => {
  try {
    const {
      waterDemand,
      intakeWell,
      pumpDesign,
      presedimentationTank,
      aerationUnit,
      rapidMix,
      clearWaterTank,
      alumDose,
      flocculatorDesign,
      gravityFilter,
      chlorinator,
    } = req.body;

    const makeHeading = (text) =>
      new Paragraph({ text, heading: HeadingLevel.HEADING_1 });

    const makeSubHeading = (text) =>
      new Paragraph({
        children: [new TextRun({ text, bold: true, break: 1 })],
      });

    const makeFormula = (formula) =>
      new Paragraph({
        children: [
          new TextRun({
            text: formula,
            font: "Courier New",
            italics: true,
          }),
        ],
      });

    // 🌊 WATER DEMAND
    let wdSection = [];
    if (waterDemand && waterDemand.Pn) {
      wdSection = [
        makeHeading("Water Demand Report"),
        makeSubHeading("📘 Formulas Used:"),
        makeFormula("Pn = P₀ × (1 + r/100)ⁿ"),
        makeFormula("WD = (Pn × per capita demand) / 1,000,000 (MLD)"),
        makeFormula("Fd = (100 × √(Pn / 1000)) / 1000 (MLD)"),
        makeFormula("Q = WD + Fd (MLD)"),
        makeFormula("q₁ = Q − 3% of Q"),
        makeFormula("q₂ = q₁ − 2% of q₁"),
        makeFormula("q₃ = q₂ − 2% of q₂"),
        makeSubHeading("📊 Results:"),
        new Paragraph({ text: `Future Population (Pn): ${waterDemand.Pn}` }),
        new Paragraph({ text: `Water Demand (WD): ${waterDemand.WD} MLD` }),
        new Paragraph({ text: `Fire Demand (Fd): ${waterDemand.Fd} MLD` }),
        new Paragraph({ text: `Total Discharge (Q): ${waterDemand.Q} MLD` }),
        new Paragraph({ text: `After 3% Loss (q1): ${waterDemand.q1} MLD` }),
        new Paragraph({ text: `After 2% Loss (q2): ${waterDemand.q2} MLD` }),
        new Paragraph({ text: `After 2% Loss (q3): ${waterDemand.q3} MLD` }),
      ];
    }

    // 💧 INTAKE WELL
    let iwSection = [];
    if (intakeWell && intakeWell.Q) {
      iwSection = [
        makeHeading("Intake Well Report"),
        makeSubHeading("📘 Formulas Used:"),
        makeFormula("Q = (Q' × 1000) / (24 × 60 × 60)"),
        makeFormula("A = Q / V"),
        makeFormula("Ah = 2 × A"),
        makeFormula("Area of One Screen = Ah / 2"),
        makeFormula("h = (Area of One Screen) / W"),
        makeSubHeading("📊 Results:"),
        new Paragraph({ text: `Discharge per Second (Q): ${intakeWell.Q} m³/sec` }),
        new Paragraph({ text: `Area for Opening (A): ${intakeWell.A} m²` }),
        new Paragraph({ text: `Total Opening Area (Ah): ${intakeWell.Ah} m²` }),
        new Paragraph({ text: `Area of One Screen: ${intakeWell.oneScreenArea} m²` }),
        new Paragraph({ text: `Height of Screen (h): ${intakeWell.h} m` }),

        new Paragraph({
          children: [
            new ImageRun({
              data: fs.readFileSync(
                path.join(__dirname, "public", "images", "img1.png")), // <-- your PNG path
              transformation: {
                width: 400,
                height: 250,
              },
            }),
          ],
        }),
        new Paragraph({ text: `Where,` }),
        new Paragraph({ text: `d1: ${intakeWell.d1} m` }),
        new Paragraph({ text: `d2: ${intakeWell.d2} m` }),
        new Paragraph({ text: `D: ${intakeWell.D} m` }),
        // new Paragraph({ text: `No. of Pipes: ${pumpDesign.H} m` })
        new Paragraph({ text: `dia: ${intakeWell.dia} m` }),
        new Paragraph({ text: `Diameter of Jackwell: ${intakeWell.D} m` })
      ];
    }



    // ⚙️ PUMP DESIGN
    let pdSection = [];
    if (pumpDesign && pumpDesign.d) {
      pdSection = [
        makeHeading("Pump Design Report"),
        makeSubHeading("📘 Formulas Used:"),
        makeFormula("d = √((4 × Q) / (π × V))"),
        makeFormula("Np = (Q × 1000) / (Pump Capacity × 86.4)"),
        makeFormula("Nt = Np + 1 (one standby pump)"),
        makeFormula("S = (0.75 × d) + 0.3"),
        makeSubHeading("📊 Results:"),
        new Paragraph({ text: `Diameter of Pipe (d): ${pumpDesign.d} m` }),
        new Paragraph({ text: `Number of Pumps (Np): ${pumpDesign.Np}` }),
        new Paragraph({ text: `Total Pumps (Nt): ${pumpDesign.Nt}` }),
        new Paragraph({ text: `Clearance Between Pumps (S): ${pumpDesign.S} m` }),
        new Paragraph({
          children: [
            new ImageRun({
              data: fs.readFileSync(path.join(__dirname, "public", "images", "img2.png")), // <-- your PNG path
              transformation: {
                width: 400,
                height: 250,
              },
            }),
          ],
        }),
        new Paragraph({ text: `Where,` }),
        new Paragraph({ text: `Nt: ${pumpDesign.Nt} m` }),
        new Paragraph({ text: `H: ${pumpDesign.H} m` }),
      ];
    }



    // 🧱 PRESEDIMENTATION TANK
    let psSection = [];
    if (presedimentationTank && presedimentationTank.V) {
      psSection = [
        makeHeading("Presedimentation Tank Report"),
        makeSubHeading("📘 Formulas Used:"),
        makeFormula("Q = (Demand × 10⁶) / (24 × 60 × 60)"),
        makeFormula("V = Q × Detention Time (m³)"),
        makeFormula("B = √(V / L)"),
        makeFormula("D = V / (L × B)"),
        makeSubHeading("📊 Results:"),
        new Paragraph({ text: `Volume (V): ${presedimentationTank.V} m³` }),
        new Paragraph({ text: `Length (L): ${presedimentationTank.L} m` }),
        new Paragraph({ text: `Width (B): ${presedimentationTank.B} m` }),
        new Paragraph({ text: `Depth (D): ${presedimentationTank.D} m` }),
        new Paragraph({
          children: [
            new ImageRun({
              data: fs.readFileSync(path.join(__dirname, "public", "images", "img3.png")), // <-- your PNG path
              transformation: {
                width: 400,
                height: 250,
              },
            }),
          ],
        }),

        new Paragraph({ text: `Where,` }),
        new Paragraph({ text: `L: ${presedimentationTank.L} m` }),
        new Paragraph({ text: `W: ${presedimentationTank.B} m` }),
      ];
    }



    // 🌬️ AERATION UNIT
    let auSection = [];
    if (aerationUnit && aerationUnit.Qp) {
      auSection = [
        makeHeading("Aeration Unit Report"),
        makeSubHeading("📘 Formulas Used:"),
        makeFormula("Q’ = (Demand × 10⁶) / 24"),
        makeFormula("A = (Q’) / (π × (Di)² / 4)"),
        makeFormula("Db = √(4 × A / π)"),
        makeFormula("t = Db / 10"),
        makeSubHeading("📊 Results:"),
        new Paragraph({ text: `Discharge per Hour (Q’): ${aerationUnit.Qp} m³/hr` }),
        new Paragraph({ text: `Inner Pipe Diameter (Di): ${aerationUnit.Di} m` }),
        new Paragraph({ text: `Tray Area (A): ${aerationUnit.A} m²` }),
        new Paragraph({ text: `Bottom Tray Diameter (Db): ${aerationUnit.Db} m` }),
        new Paragraph({ text: `Tray Tread (t): ${aerationUnit.t} m` }),
        new Paragraph({
          children: [
            new ImageRun({
              data: fs.readFileSync(path.join(__dirname, "public", "images", "img4.png")), // <-- your PNG path
              transformation: {
                width: 400,
                height: 250,
              },
            }),
          ],
        }),

        new Paragraph({ text: `Where,` }),
        new Paragraph({ text: `Db: ${aerationUnit.Db} m` }),
      ];
    }



    // ⚡ RAPID MIX
    let rmSection = [];
    if (rapidMix && rapidMix.Qp) {
      rmSection = [
        makeHeading("Rapid Mix Report"),
        makeSubHeading("📘 Formulas Used:"),
        makeFormula("Q’ = (Q × 10⁶) / 24"),
        makeFormula("C = Q’ × Detention Time / 60"),
        makeFormula("D = √(4 × C / (π × H))"),
        makeFormula("HP = (P × N × 9.81 × 10⁻³) / Efficiency"),
        makeSubHeading("📊 Results:"),
        new Paragraph({ text: `Design Flow (Q’): ${rapidMix.Qp} m³/hr` }),
        new Paragraph({ text: `Tank Capacity (C): ${rapidMix.C} m³` }),
        new Paragraph({ text: `Tank Diameter (D): ${rapidMix.D} m` }),
        new Paragraph({ text: `Tank Volume (V): ${rapidMix.V} m³` }),
        new Paragraph({ text: `No. of Units: ${rapidMix.no}` }),
        new Paragraph({ text: `Motor Power (HP): ${rapidMix.HP} HP` }),
        new Paragraph({ text: `Impeller Diameter (d): ${rapidMix.d} m` }),
        new Paragraph({
          children: [
            new ImageRun({
              data: fs.readFileSync(path.join(__dirname, "public", "images", "img5.png")), // <-- your PNG path
              transformation: {
                width: 400,
                height: 250,
              },
            }),
          ],
        }),
        new Paragraph({ text: `Where,` }),
        new Paragraph({ text: `D: ${rapidMix.D} m` }),
        new Paragraph({ text: `H: ${rapidMix.H} m` }),
      ];
    }

    // 🧪 ALUM DOSE
    let ads = [];
    if (alumDose && alumDose.R) {
      ads = [
        makeHeading("Alum Dose Report"),
        new Paragraph({ text: `Alum Required per Hour (R): ${alumDose.R} g/hr` }),
        new Paragraph({ text: `Per Day (W): ${alumDose.W} kg/day` }),
        new Paragraph({ text: `For n months (Wt): ${alumDose.Wt} kg` }),
        new Paragraph({ text: `Tank Volume (V1): ${alumDose.V1} m³` }),
        new Paragraph({ text: `Provision Volume (V2): ${alumDose.V2} m³` }),
        new Paragraph({ text: `Total Volume (V): ${alumDose.V} m³` }),
        new Paragraph({ text: `Tank Diameter: ${alumDose.dia} m` }),
        new Paragraph({ text: `Square Platform Side (l): ${alumDose.l} m` }),
      ];
    }

    // 🔄 FLOCCULATOR DESIGN
    let fds = [];
    if (flocculatorDesign && flocculatorDesign.Q) {
      fds = [
        makeHeading("Flocculator Design Report"),
        new Paragraph({ text: `Outflow (Q): ${flocculatorDesign.Q} m³/hr` }),
        new Paragraph({ text: `Flocculator Volume (V): ${flocculatorDesign.V} m³` }),
        new Paragraph({ text: `Plan Area (A): ${flocculatorDesign.A} m²` }),
        new Paragraph({ text: `Diameter (D): ${flocculatorDesign.D} m` }),
        new Paragraph({ text: "Clarifier", heading: HeadingLevel.HEADING_3 }),
        new Paragraph({ text: `Clarifier Surface Area (Ac): ${flocculatorDesign.Ac} m²` }),
        new Paragraph({ text: `Clariflocculator Diameter (D’): ${flocculatorDesign.Dp} m` }),
        new Paragraph({ text: `Weir Length (L): ${flocculatorDesign.L} m` }),
        new Paragraph({ text: `Weir Loading (F): ${flocculatorDesign.F} m³/m·day` }),
        new Paragraph({ text: `Tank Depth (d): ${flocculatorDesign.d} m` }),
        new Paragraph({ text: `Sludge Depth (d1): ${flocculatorDesign.d1} m` }),
        new Paragraph({ text: `Total Depth (d’): ${flocculatorDesign.dtotal} m` }),
        new Paragraph({ text: "Paddles", heading: HeadingLevel.HEADING_3 }),
        new Paragraph({ text: `Paddle Area (Ap): ${flocculatorDesign.Ap_calc} m²` }),
        new Paragraph({ text: `Paddle Area (a): ${flocculatorDesign.a} m²` }),
        new Paragraph({ text: `Shaft Distance (s): ${flocculatorDesign.s} m` }),
        new Paragraph({ text: `Total Paddles (Tno): ${flocculatorDesign.Tno}` }),
        new Paragraph({ text: "Launder", heading: HeadingLevel.HEADING_3 }),
        new Paragraph({ text: `Flow (q): ${flocculatorDesign.q} m³/hr` }),
        new Paragraph({ text: `Launder Area (a’): ${flocculatorDesign.aL} m²` }),
        new Paragraph({ text: `Perimeter (P): ${flocculatorDesign.Pperimeter} m` }),
        new Paragraph({ text: `Mean Radius (R): ${flocculatorDesign.Rm}` }),
        new Paragraph({ text: `Slope (S): ${flocculatorDesign.S}` }),
      ];
    }

    // 🧱 GRAVITY FILTER
    let gf = [];
    if (gravityFilter && gravityFilter.Q1) {
      gf = [
        makeHeading("Gravity Filter Report"),
        new Paragraph({ text: `Design Flow (Q1): ${gravityFilter.Q1} m³/day` }),
        new Paragraph({ text: `Filter Area (A): ${gravityFilter.A} m²` }),
        new Paragraph({ text: `No. of Filters: ${gravityFilter.no}` }),
        new Paragraph({ text: `Area Each (A’): ${gravityFilter.A1} m²` }),
        new Paragraph({ text: `Total Perforation Area: ${gravityFilter.a} m²` }),
        new Paragraph({ text: `Total No. of Perforations: ${gravityFilter.num}` }),
        new Paragraph({ text: `Manifold Diameter (Qm): ${gravityFilter.Qm} m` }),
        new Paragraph({ text: `Laterals on Both Sides (Nbl): ${gravityFilter.Nbl}` }),
        new Paragraph({ text: `Total Tanks (No_Tank): ${gravityFilter.No_Tank}` }),
      ];
    }

    // 🧂 CHLORINATOR
    let chl = [];
    if (chlorinator && chlorinator.totalChlorineApplied) {
      chl = [
        makeHeading("Chlorinator Report"),
        new Paragraph({ text: `Total Chlorine Applied: ${chlorinator.totalChlorineApplied} mg/h` }),
        new Paragraph({ text: `Chlorine per Hour (R): ${chlorinator.R} mg/day` }),
        new Paragraph({ text: `Chlorine per Day (W): ${chlorinator.W} mg` }),
        new Paragraph({ text: `Total Chlorine Required (Wt): ${chlorinator.Wt} m³` }),
        new Paragraph({ text: `Tank Volume (V1): ${chlorinator.V1} m³` }),
        new Paragraph({ text: `Mixing Volume (V2): ${chlorinator.V2} m³` }),
        new Paragraph({ text: `Total Volume (V): ${chlorinator.totalVolume} m³` }),
        new Paragraph({ text: `Tank Diameter: ${chlorinator.Dia} m` }),
        new Paragraph({ text: `Square Platform (l): ${chlorinator.l} m` }),
      ];
    }

    // 🏗️ CLEAR WATER TANK (final revised)
    let cwtSection = [];
    if (clearWaterTank && clearWaterTank.A) {
      cwtSection = [
        makeHeading("Clear Water Tank Report"),
        new Paragraph({ text: `Cross Sectional Area (A): ${clearWaterTank.A} m²` }),
        new Paragraph({ text: `Diameter (d): ${clearWaterTank.diameter} m` }),
      ];
    }

    // ✅ Combine Everything
    const doc = new Document({
      sections: [
        {
          children: [
            ...wdSection,
            ...iwSection,
            ...pdSection,
            ...psSection,
            ...auSection,
            ...rmSection,
            ...ads,
            ...fds,
            ...gf,
            ...chl,
            ...cwtSection,
          ],
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader("Content-Disposition", "attachment; filename=ProjectReport.docx");
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.send(buffer);
  } catch (err) {
    console.error("Error:", err);
    res.status(500).send("Error generating report");
  }
});

app.listen(5000, () => console.log("✅ Server running on http://localhost:5000"));
