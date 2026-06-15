const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Eric Kharitonov";
pres.title = "Deutsch-Jozsa Algorithm Replication";

// ── PALETTE ─────────────────────────────────────────────────────────────────
const BG      = "060F1D";   // slide background – very dark navy
const CODE_BG = "0C1B2E";   // code block fill
const CODE_BD = "1C3A5C";   // code block border
const CODE_TX = "B8D8ED";   // code text
const ACC     = "00AACC";   // cyan accent
const ACC2    = "00425A";   // dark accent background
const WHT     = "FFFFFF";
const LBL     = "8CBBD0";   // body text
const MUT     = "456070";   // muted text

// ── REUSABLE HELPERS ────────────────────────────────────────────────────────

function codeBox(slide, code, x, y, w, h) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: CODE_BG },
    line: { color: CODE_BD, width: 1 },
  });
  slide.addText(code, {
    x: x + 0.18, y: y + 0.12, w: w - 0.36, h: h - 0.24,
    fontFace: "Consolas", fontSize: 11, color: CODE_TX,
    valign: "top", margin: 0,
  });
}

function slideTitle(slide, t, sub) {
  slide.addText(t, {
    x: 0.5, y: 0.15, w: 9.0, h: 0.62,
    fontFace: "Calibri", fontSize: 30, bold: true, color: WHT, margin: 0,
  });
  if (sub) {
    slide.addText(sub, {
      x: 0.5, y: 0.77, w: 9.0, h: 0.28,
      fontFace: "Calibri", fontSize: 13, color: ACC, italic: true, margin: 0,
    });
  }
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.07, w: 9.0, h: 0.02,
    fill: { color: ACC2 }, line: { color: ACC2 },
  });
}

function infoCard(slide, label, desc, x, y, w, h) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: CODE_BG },
    line: { color: CODE_BD, width: 1 },
  });
  slide.addText(label, {
    x: x + 0.14, y: y + 0.08, w: w - 0.28, h: 0.27,
    fontFace: "Consolas", fontSize: 11, bold: true, color: ACC, margin: 0,
  });
  slide.addText(desc, {
    x: x + 0.14, y: y + 0.36, w: w - 0.28, h: h - 0.44,
    fontFace: "Calibri", fontSize: 11, color: LBL, margin: 0,
  });
}

// ── SLIDE 1 : TITLE ─────────────────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.14, h: 5.625,
    fill: { color: ACC }, line: { color: ACC },
  });
  s.addText("Deutsch-Jozsa", {
    x: 0.5, y: 1.1, w: 9.0, h: 1.1,
    fontFace: "Calibri", fontSize: 62, bold: true, color: WHT, margin: 0,
  });
  s.addText("Algorithm Replication", {
    x: 0.5, y: 2.2, w: 9.0, h: 0.7,
    fontFace: "Calibri", fontSize: 32, color: ACC, margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.0, w: 3.5, h: 0.04,
    fill: { color: ACC2 }, line: { color: ACC2 },
  });
  s.addText("Eric Kharitonov  ·  Qiskit + IBM Quantum", {
    x: 0.5, y: 3.15, w: 9.0, h: 0.35,
    fontFace: "Calibri", fontSize: 16, color: MUT, margin: 0,
  });

  // Decorative circuit diagram (top-right)
  const gx = [7.5, 8.1, 8.7];
  const gy = [1.6, 2.2, 2.8];
  const gLabels = [["H","U","H"],["H","U","H"],["H","U",""]];
  for (let row = 0; row < 3; row++) {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 7.1, y: gy[row] + 0.14, w: 2.1, h: 0.02,
      fill: { color: ACC2 }, line: { color: ACC2 },
    });
    for (let col = 0; col < 3; col++) {
      if (!gLabels[row][col]) continue;
      s.addShape(pres.shapes.RECTANGLE, {
        x: gx[col] - 0.17, y: gy[row], w: 0.34, h: 0.30,
        fill: { color: ACC2 }, line: { color: ACC, width: 1 },
      });
      s.addText(gLabels[row][col], {
        x: gx[col] - 0.17, y: gy[row], w: 0.34, h: 0.30,
        fontFace: "Calibri", fontSize: 10, bold: true, color: ACC,
        align: "center", valign: "middle", margin: 0,
      });
    }
  }
}

// ── SLIDE 2 : THE PROBLEM ───────────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slideTitle(s, "The Problem", "Constant or Balanced — and how do you know?");

  s.addText("You have a black-box function f that returns 0 or 1.\nYou are told it is one of two types:", {
    x: 0.5, y: 1.22, w: 5.2, h: 0.65,
    fontFace: "Calibri", fontSize: 14, color: LBL, margin: 0,
  });

  // Constant card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 2.0, w: 5.2, h: 0.9,
    fill: { color: CODE_BG }, line: { color: ACC, width: 1 },
  });
  s.addText("CONSTANT", {
    x: 0.5, y: 2.05, w: 5.2, h: 0.38,
    fontFace: "Calibri", fontSize: 17, bold: true, color: ACC,
    align: "center", margin: 0,
  });
  s.addText("Returns the same value (0 or 1) for every input", {
    x: 0.5, y: 2.44, w: 5.2, h: 0.35,
    fontFace: "Calibri", fontSize: 13, color: LBL,
    align: "center", margin: 0,
  });

  // Balanced card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.05, w: 5.2, h: 0.9,
    fill: { color: CODE_BG }, line: { color: MUT, width: 1 },
  });
  s.addText("BALANCED", {
    x: 0.5, y: 3.1, w: 5.2, h: 0.38,
    fontFace: "Calibri", fontSize: 17, bold: true, color: LBL,
    align: "center", margin: 0,
  });
  s.addText("Returns 0 for exactly half the inputs, 1 for the other half", {
    x: 0.5, y: 3.49, w: 5.2, h: 0.35,
    fontFace: "Calibri", fontSize: 13, color: LBL,
    align: "center", margin: 0,
  });

  // Right panel: Quantum advantage
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.0, y: 1.22, w: 3.5, h: 3.73,
    fill: { color: CODE_BG }, line: { color: CODE_BD, width: 1 },
  });
  s.addText("Classical", {
    x: 6.15, y: 1.33, w: 3.2, h: 0.38,
    fontFace: "Calibri", fontSize: 16, bold: true, color: MUT, margin: 0,
  });
  s.addText("Worst case: must check more\nthan half of all possible inputs", {
    x: 6.15, y: 1.73, w: 3.2, h: 0.7,
    fontFace: "Calibri", fontSize: 13, color: LBL, margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.15, y: 2.55, w: 3.1, h: 0.02,
    fill: { color: ACC2 }, line: { color: ACC2 },
  });
  s.addText("Quantum", {
    x: 6.15, y: 2.67, w: 3.2, h: 0.38,
    fontFace: "Calibri", fontSize: 16, bold: true, color: ACC, margin: 0,
  });
  s.addText("1 query.\nEvery time.", {
    x: 6.15, y: 3.05, w: 3.2, h: 0.72,
    fontFace: "Calibri", fontSize: 24, bold: true, color: WHT, margin: 0,
  });
  s.addText("That's the promise of Deutsch-Jozsa.", {
    x: 6.15, y: 3.82, w: 3.2, h: 0.45,
    fontFace: "Calibri", fontSize: 12, color: MUT, italic: true, margin: 0,
  });
}

// ── SLIDE 3 : SETUP & CONNECTION ────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slideTitle(s, "1. Setup & Connection", "Imports · Backend · Noise Model");

  const code3 =
`from qiskit_ibm_runtime import QiskitRuntimeService
from qiskit_ibm_runtime import SamplerV2 as Sampler
from qiskit.primitives import BackendSamplerV2
from qiskit_aer import AerSimulator
from qiskit_aer.noise import NoiseModel
from qiskit import QuantumCircuit
import numpy as np

# Connect to the least-busy real IBM backend
service = QiskitRuntimeService()
backend = service.least_busy(
    operational=True, simulator=False, min_num_qubits=6
)
# Mirror that backend's noise into a local simulator
noise_model = NoiseModel.from_backend(backend)
backend_sim = AerSimulator(noise_model=noise_model)
sampler_sim = BackendSamplerV2(backend=backend_sim)`;

  codeBox(s, code3, 0.5, 1.15, 5.5, 3.85);

  const cards3 = [
    ["least_busy()",    "Selects the live IBM device with the shortest job queue"],
    ["NoiseModel",      "Captures real hardware error rates for realistic simulation"],
    ["AerSimulator",    "Runs the noisy circuit locally — no hardware queue needed"],
    ["BackendSamplerV2","Executes transpiled circuits and collects measurement counts"],
  ];
  let cy = 1.15;
  for (const [lbl, dsc] of cards3) {
    infoCard(s, lbl, dsc, 6.3, cy, 3.2, 0.87);
    cy += 0.97;
  }
}

// ── SLIDE 4 : ORACLE — CONSTANT ─────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slideTitle(s, "2. The Oracle", "Constant Function");

  const code4 =
`def dj_function(num_qubits):
    qc_dj = QuantumCircuit(num_qubits + 1)

    # Maybe flip the ancilla (still a constant function)
    if np.random.randint(0, 2):
        qc_dj.x(num_qubits)

    # 50% chance: return early -> constant function
    if np.random.randint(0, 2):
        return qc_dj

    # Otherwise: continue to build balanced...`;

  codeBox(s, code4, 0.5, 1.15, 5.5, 2.8);

  // Right explanation
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.3, y: 1.15, w: 3.2, h: 2.8,
    fill: { color: CODE_BG }, line: { color: CODE_BD, width: 1 },
  });
  s.addText("Constant oracle", {
    x: 6.44, y: 1.26, w: 2.92, h: 0.30,
    fontFace: "Calibri", fontSize: 14, bold: true, color: ACC, margin: 0,
  });
  s.addText(
    "The function returns the same value for every input.\n\n" +
    "The ancilla flip only swaps constant-0 to constant-1 — still constant.\n\n" +
    "If the second coin flip returns 1 we exit early. The oracle does nothing to the input qubits. That is the constant case.",
    {
      x: 6.44, y: 1.60, w: 2.92, h: 2.2,
      fontFace: "Calibri", fontSize: 12, color: LBL, margin: 0,
    }
  );

  // Example row
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.1, w: 5.5, h: 0.7,
    fill: { color: ACC2 }, line: { color: CODE_BD, width: 1 },
  });
  s.addText("f(00000)=0   f(00001)=0   f(10101)=0   ...   f(11111)=0", {
    x: 0.65, y: 4.1, w: 5.2, h: 0.7,
    fontFace: "Consolas", fontSize: 12, color: CODE_TX,
    valign: "middle", margin: 0,
  });
}

// ── SLIDE 5 : ORACLE — BALANCED ─────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slideTitle(s, "2. The Oracle", "Balanced Function — The Conjugate Pattern");

  const code5 =
`on_states = np.random.choice(
    range(2**num_qubits),
    2**num_qubits // 2,   # exactly half of all states
    replace=False,
)

def add_cx(qc_dj, bit_string):
    for qubit, bit in enumerate(reversed(bit_string)):
        if bit == "1":
            qc_dj.x(qubit)
    return qc_dj

for state in on_states:
    qc_dj = add_cx(qc_dj, f"{state:0{num_qubits}b}")
    qc_dj.mcx(list(range(num_qubits)), num_qubits)
    qc_dj = add_cx(qc_dj, f"{state:0{num_qubits}b}")`;

  codeBox(s, code5, 0.5, 1.15, 5.5, 3.85);

  const steps5 = [
    ["X gates ON",  "Flip qubits so this state looks like |11111>"],
    ["mcx(...)",    "Multi-controlled X fires — ancilla flips for this state only"],
    ["X gates OFF", "Restore qubits (same flips undo each other)"],
  ];
  let sy = 1.15;
  for (const [lbl, dsc] of steps5) {
    infoCard(s, lbl, dsc, 6.3, sy, 3.2, 1.05);
    sy += 1.15;
  }

  s.addText("Half the states flip the ancilla. Half do not. That is what makes it balanced.", {
    x: 0.5, y: 5.1, w: 5.5, h: 0.35,
    fontFace: "Calibri", fontSize: 12, italic: true, color: MUT, margin: 0,
  });
}

// ── SLIDE 6 : BUILD THE CIRCUIT ─────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slideTitle(s, "3. Build the Circuit", "6 lines. The complete Deutsch-Jozsa circuit.");

  const code6 =
`n = 5  # number of input qubits

qc_dj = QuantumCircuit(n + 1, n)

qc_dj.x(n)                           # ancilla to |1>
qc_dj.h(range(n + 1))                # Hadamard all qubits
qc_dj.barrier()
qc_dj.compose(oracle, inplace=True)  # apply oracle (phase kickback)
qc_dj.barrier()
qc_dj.h(range(n))                    # Hadamard inputs only
qc_dj.measure(range(n), range(n))    # measure`;

  codeBox(s, code6, 0.5, 1.15, 5.5, 2.95);

  const steps6 = [
    ["x(n)",       "Ancilla to |1>  (becomes |-> after Hadamard)"],
    ["h(all)",     "Inputs into superposition. Ancilla into |->"],
    ["oracle",     "Phase kickback encodes f into the input phases"],
    ["h(inputs)",  "Interference converts phases into measurable bits"],
    ["measure",    "Read the answer from the input register"],
  ];
  let sy6 = 1.15;
  for (const [lbl, dsc] of steps6) {
    infoCard(s, lbl, dsc, 6.3, sy6, 3.2, 0.76);
    sy6 += 0.84;
  }

  // Simple circuit sketch at bottom
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.27, w: 5.5, h: 1.07,
    fill: { color: ACC2 }, line: { color: CODE_BD, width: 1 },
  });
  s.addText(
    "|0> -- H -- [  oracle  ] -- H -- M\n" +
    "|0> -- H -- [  oracle  ] -- H -- M\n" +
    "|1> -- H -- [  oracle  ] ---------",
    {
      x: 0.65, y: 4.27, w: 5.2, h: 1.07,
      fontFace: "Consolas", fontSize: 11, color: CODE_TX,
      valign: "middle", margin: 0,
    }
  );
}

// ── SLIDE 7 : TRANSPILE ──────────────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slideTitle(s, "4. Transpile for Hardware", "Rewrite the circuit in the hardware's native gate language");

  const code7 =
`target = backend.target
pm = generate_preset_pass_manager(
    target=target,
    optimization_level=3
)
qc_isa = pm.run(qc_dj)`;

  codeBox(s, code7, 0.5, 1.15, 5.5, 1.8);

  // Before/After stats
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.1, w: 5.5, h: 1.15,
    fill: { color: CODE_BG }, line: { color: CODE_BD, width: 1 },
  });
  s.addText("Original circuit:    depth 4,  19 gates", {
    x: 0.68, y: 3.2, w: 5.1, h: 0.38,
    fontFace: "Consolas", fontSize: 13, color: MUT, margin: 0,
  });
  s.addText("Transpiled circuit:  depth 7,  40 gates", {
    x: 0.68, y: 3.6, w: 5.1, h: 0.38,
    fontFace: "Consolas", fontSize: 13, color: ACC, margin: 0,
  });

  // Right explanation
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.3, y: 1.15, w: 3.2, h: 3.1,
    fill: { color: CODE_BG }, line: { color: CODE_BD, width: 1 },
  });
  s.addText("Why transpile?", {
    x: 6.44, y: 1.26, w: 2.92, h: 0.30,
    fontFace: "Calibri", fontSize: 14, bold: true, color: ACC, margin: 0,
  });
  s.addText(
    "Real quantum hardware supports only a small set of native gates " +
    "(CX, Rz, SX) on specific qubit pairs.\n\n" +
    "Transpilation rewrites your circuit into those native gates and routes " +
    "operations across the chip's physical qubit topology.\n\n" +
    "More gates, same result — the hardware just needs smaller steps.",
    {
      x: 6.44, y: 1.60, w: 2.92, h: 2.4,
      fontFace: "Calibri", fontSize: 12, color: LBL, margin: 0,
    }
  );

  // Backend label
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.4, w: 5.5, h: 0.7,
    fill: { color: ACC2 }, line: { color: CODE_BD, width: 1 },
  });
  s.addText("Backend:  ibm_fez   |   optimization_level: 3", {
    x: 0.65, y: 4.4, w: 5.2, h: 0.7,
    fontFace: "Consolas", fontSize: 13, bold: true, color: ACC,
    valign: "middle", margin: 0,
  });
}

// ── SLIDE 8 : RUN THE CIRCUIT ────────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slideTitle(s, "5. Run the Circuit", "shots = 1");

  const code8 =
`USE_SIMULATOR = True

if USE_SIMULATOR:
    job = sampler_sim.run([qc_isa], shots=1)
else:
    job = sampler.run([qc_isa], shots=1)

print(f"Job ID: {job.job_id()}")

while True:
    status = job.status()
    print(f"Status: {status}")
    if 'DONE' in str(status):
        break
    time.sleep(10)

print("Job completed!")`;

  codeBox(s, code8, 0.5, 1.15, 5.5, 3.5);

  // Right explanation
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.3, y: 1.15, w: 3.2, h: 3.5,
    fill: { color: CODE_BG }, line: { color: CODE_BD, width: 1 },
  });
  s.addText("Why shots = 1?", {
    x: 6.44, y: 1.26, w: 2.92, h: 0.30,
    fontFace: "Calibri", fontSize: 14, bold: true, color: ACC, margin: 0,
  });
  s.addText(
    "Deutsch-Jozsa is deterministic. " +
    "Quantum interference guarantees the same correct answer on every single run.\n\n" +
    "One shot is sufficient.\n\n" +
    "The USE_SIMULATOR flag lets you switch between the noisy local simulator " +
    "and real IBM hardware without touching the logic.",
    {
      x: 6.44, y: 1.60, w: 2.92, h: 2.9,
      fontFace: "Calibri", fontSize: 12, color: LBL, margin: 0,
    }
  );

  // Console output strip
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.8, w: 5.5, h: 0.55,
    fill: { color: CODE_BG }, line: { color: CODE_BD, width: 1 },
  });
  s.addText("Status: JobStatus.RUNNING  =>  Status: JobStatus.DONE  =>  Job completed!", {
    x: 0.65, y: 4.8, w: 5.2, h: 0.55,
    fontFace: "Consolas", fontSize: 10, color: CODE_TX,
    valign: "middle", margin: 0,
  });
}

// ── SLIDE 9 : INTERPRET RESULTS ──────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slideTitle(s, "5. Read the Results", "All zeros = constant.  Any ones = balanced.");

  const code9 =
`res = job.result()
counts = res[0].data.c.get_counts()

print(f"Measurement result: {counts}")

if "0" * n in counts:
    print("Result: CONSTANT function")
else:
    print("Result: BALANCED function")`;

  codeBox(s, code9, 0.5, 1.15, 5.5, 2.2);

  // Decision visual
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.52, w: 2.5, h: 1.55,
    fill: { color: ACC2 }, line: { color: ACC, width: 2 },
  });
  s.addText("00000", {
    x: 0.5, y: 3.62, w: 2.5, h: 0.6,
    fontFace: "Consolas", fontSize: 22, bold: true, color: WHT,
    align: "center", margin: 0,
  });
  s.addText("CONSTANT", {
    x: 0.5, y: 4.24, w: 2.5, h: 0.42,
    fontFace: "Calibri", fontSize: 14, bold: true, color: ACC,
    align: "center", margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 3.25, y: 3.52, w: 2.5, h: 1.55,
    fill: { color: CODE_BG }, line: { color: MUT, width: 1 },
  });
  s.addText("01010", {
    x: 3.25, y: 3.62, w: 2.5, h: 0.6,
    fontFace: "Consolas", fontSize: 22, bold: true, color: MUT,
    align: "center", margin: 0,
  });
  s.addText("BALANCED", {
    x: 3.25, y: 4.24, w: 2.5, h: 0.42,
    fontFace: "Calibri", fontSize: 14, bold: true, color: MUT,
    align: "center", margin: 0,
  });

  // Right explanation
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.3, y: 1.15, w: 3.2, h: 3.92,
    fill: { color: CODE_BG }, line: { color: CODE_BD, width: 1 },
  });
  s.addText("Why does it work?", {
    x: 6.44, y: 1.26, w: 2.92, h: 0.30,
    fontFace: "Calibri", fontSize: 14, bold: true, color: ACC, margin: 0,
  });
  s.addText(
    "Constant function: all inputs pick up the same phase. " +
    "The second Hadamard layer constructively interferes everything back into |00000>.\n\n" +
    "Balanced function: inputs pick up opposite phases. " +
    "Interference cancels the |00000> amplitude entirely. You always get a non-zero result.\n\n",
    {
      x: 6.44, y: 1.60, w: 2.92, h: 2.6,
      fontFace: "Calibri", fontSize: 12, color: LBL, margin: 0,
    }
  );
  s.addText("One query. Certain answer.", {
    x: 6.44, y: 4.18, w: 2.92, h: 0.45,
    fontFace: "Calibri", fontSize: 14, bold: true, italic: true, color: ACC, margin: 0,
  });
}

// ── SLIDE 10 : FULL OUTPUT ───────────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };
  slideTitle(s, "Full Output", "");

  const output =
`Using backend: ibm_fez
Oracle created with 6 qubits
Original circuit:    depth 4,  19 gates
Transpiled circuit:  depth 7,  40 gates

Running on simulator...

Job ID: ed128622-ef3c-4637-9a73-100421cbf125
Waiting for job to complete...
Status: JobStatus.RUNNING
Status: JobStatus.DONE

Job completed!

Measurement result: {'00000': 1}
Result: CONSTANT function`;

  codeBox(s, output, 1.2, 1.15, 7.6, 3.85);

  // Highlight bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 1.2, y: 4.65, w: 7.6, h: 0.45,
    fill: { color: ACC2 }, line: { color: ACC, width: 1 },
  });
  s.addText("{'00000': 1}  -->  all zeros  -->  CONSTANT  ✓", {
    x: 1.2, y: 4.65, w: 7.6, h: 0.45,
    fontFace: "Consolas", fontSize: 14, bold: true, color: ACC,
    align: "center", valign: "middle", margin: 0,
  });
}

// ── SLIDE 11 : CONCLUSION ────────────────────────────────────────────────────
{
  let s = pres.addSlide();
  s.background = { color: BG };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.14, h: 5.625,
    fill: { color: ACC }, line: { color: ACC },
  });

  s.addText("That's Deutsch-Jozsa.", {
    x: 0.5, y: 0.65, w: 9.0, h: 0.85,
    fontFace: "Calibri", fontSize: 44, bold: true, color: WHT, margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.57, w: 4.0, h: 0.04,
    fill: { color: ACC2 }, line: { color: ACC2 },
  });

  const bullets = [
    ["Random oracle — constant or balanced",              false],
    ["Ancilla to |-> enables phase kickback",             false],
    ["Superposition evaluates all inputs simultaneously", false],
    ["Interference collapses phases into measurable bits",false],
    ["One query. Certain answer.",                        true ],
  ];
  let by = 1.75;
  for (const [text, highlight] of bullets) {
    s.addText(text, {
      x: 0.7, y: by, w: 5.6, h: 0.42,
      fontFace: "Calibri", fontSize: highlight ? 16 : 14,
      bold: highlight, italic: highlight,
      color: highlight ? ACC : LBL,
      bullet: true, margin: 0,
    });
    by += 0.45;
  }

  // Section recap card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.8, y: 1.57, w: 2.7, h: 2.65,
    fill: { color: CODE_BG }, line: { color: CODE_BD, width: 1 },
  });
  s.addText("Code Sections", {
    x: 6.94, y: 1.68, w: 2.42, h: 0.30,
    fontFace: "Calibri", fontSize: 13, bold: true, color: ACC, margin: 0,
  });
  const secs = [
    "1. Setup & Connection",
    "2. Oracle Function",
    "3. Build Circuit",
    "4. Transpile",
    "5. Run & Interpret",
  ];
  let secY = 2.05;
  for (const sec of secs) {
    s.addText(sec, {
      x: 6.94, y: secY, w: 2.42, h: 0.35,
      fontFace: "Calibri", fontSize: 12, color: LBL, margin: 0,
    });
    secY += 0.39;
  }

  // GitHub footer
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.88, w: 9.0, h: 0.55,
    fill: { color: ACC2 }, line: { color: CODE_BD, width: 1 },
  });
  s.addText("github.com/eric-kharitonov/Deutsch-Jozsa-Replication", {
    x: 0.65, y: 4.88, w: 8.7, h: 0.55,
    fontFace: "Calibri", fontSize: 14, color: CODE_TX,
    valign: "middle", margin: 0,
  });
}

// ── WRITE FILE ───────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "C:\\work\\deutsch-jozsa-slides.pptx" })
  .then(() => console.log("Saved: C:\\work\\deutsch-jozsa-slides.pptx"))
  .catch(err => console.error("Error:", err));
