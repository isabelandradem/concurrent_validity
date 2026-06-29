/* Khan Academy Kids — Boston Raffle picker */
(() => {
  "use strict";

  // ----- Configuration -----
  // Only sample people tagged with this round. Set to "" or null to disable filtering.
  const ROUND_FILTER = "Round 1";

  // ----- Bundled attendees fallback (Round 1 from Boston Gates (Responses).xlsx) -----
  // Used when the page is opened via file:// where fetch() can't load the .xlsx.
  // The user can also upload a fresh spreadsheet to override this.
  const FALLBACK_ATTENDEES = [
  {
    "name": "Heather Adams",
    "organization": "Guilford Cty Partnership for Children"
  },
  {
    "name": "Amanda Mason",
    "organization": "Tennessee Department of education"
  },
  {
    "name": "Kristie Kylander",
    "organization": "Guilford County Partnership for Children"
  },
  {
    "name": "Sommer Wynn",
    "organization": "KIPP DC"
  },
  {
    "name": "Tommy Welch",
    "organization": "Boston Public Schools"
  },
  {
    "name": "Y",
    "organization": "Sacramento County Office Of Education"
  },
  {
    "name": "Dr. Nakeisha Savage",
    "organization": "Maryland State Department of Education"
  },
  {
    "name": "Abigail Daniels",
    "organization": "Dc public schools"
  },
  {
    "name": "Silvia Abbato",
    "organization": "Union City Board of Education"
  },
  {
    "name": "Crystal Eldridge",
    "organization": "CCSSO"
  },
  {
    "name": "Gizelle McIntyre",
    "organization": "DCPS"
  },
  {
    "name": "Willy Wong",
    "organization": "Ars Marinus"
  },
  {
    "name": "Lauren Chisholm",
    "organization": "District of Columbia Public Schools"
  },
  {
    "name": "Tyson",
    "organization": "Praidea"
  },
  {
    "name": "claudine milano",
    "organization": "Boston Public Schools"
  },
  {
    "name": "TeeAra Dias",
    "organization": "Boston Public Schools"
  },
  {
    "name": "Tom Green",
    "organization": "EdReports"
  },
  {
    "name": "Rosalyn Rice-Harris",
    "organization": "CCSSO"
  },
  {
    "name": "Carolyn Christopher",
    "organization": "Boston public schools"
  },
  {
    "name": "Sheila Briggs",
    "organization": "Education First"
  },
  {
    "name": "Mayra Cuevas",
    "organization": "Boston Public Schools"
  },
  {
    "name": "Juan Jiménez",
    "organization": "Education First"
  },
  {
    "name": "Jade",
    "organization": "University of California Irvine"
  }
];

  // ----- DOM refs -----
  const reel = document.querySelector(".reel");
  const reelName = document.getElementById("reelName");
  const reelOrg = document.getElementById("reelOrg");
  const pickBtn = document.getElementById("pickBtn");
  const resetBtn = document.getElementById("resetBtn");
  const fileInput = document.getElementById("fileInput");
  const countNum = document.getElementById("countNum");
  const hint = document.getElementById("hint");
  const overlay = document.getElementById("winnerOverlay");
  const winnerName = document.getElementById("winnerName");
  const winnerOrg = document.getElementById("winnerOrg");
  const closeWinner = document.getElementById("closeWinner");
  const spinAgain = document.getElementById("spinAgain");

  // ----- State -----
  let attendees = [];        // [{name, organization}]
  let isSpinning = false;

  // ----- Helpers -----
  const setHint = (msg, isError = false) => {
    hint.innerHTML = msg;
    hint.classList.toggle("error", !!isError);
  };

  const setAttendees = (list, source, totalBeforeFilter) => {
    attendees = list.filter((a) => a && a.name);
    countNum.textContent = String(attendees.length);
    const filterTag = ROUND_FILTER ? ` <strong>${ROUND_FILTER}</strong>` : "";
    const totalNote =
      typeof totalBeforeFilter === "number" && totalBeforeFilter !== attendees.length
        ? ` (filtered from ${totalBeforeFilter})`
        : "";
    if (attendees.length) {
      setHint(
        `Loaded <strong>${attendees.length}</strong>${filterTag} attendees from <code>${source}</code>${totalNote}. Ready to pick a winner!`,
      );
      pickBtn.disabled = false;
    } else {
      setHint(
        `No${filterTag} attendees found in <code>${source}</code>. Try uploading a different file.`,
        true,
      );
      pickBtn.disabled = true;
    }
  };

  // Parse an xlsx/csv ArrayBuffer using SheetJS, finding Name + Organization (+ Round) columns
  // Returns { rows: [...], totalBeforeFilter }
  const parseSpreadsheet = (arrayBuffer) => {
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
    if (!rows.length) return { rows: [], totalBeforeFilter: 0 };

    const keys = Object.keys(rows[0]);
    const findKey = (patterns) =>
      keys.find((k) =>
        patterns.some((p) => k.toLowerCase().trim().includes(p)),
      );
    const nameKey = findKey(["name"]) || keys[0];
    const orgKey =
      findKey(["organization", "organisation", "company", "org"]) ||
      keys[1] ||
      "";

    // Round tag column: explicit "round" header, the form's "Column 1", or any column
    // whose first non-empty value starts with "Round "
    let roundKey = findKey(["round", "group", "tag"]);
    if (!roundKey && keys.includes("Column 1")) roundKey = "Column 1";
    if (!roundKey) {
      roundKey = keys.find((k) =>
        rows.some((r) => /^round\s*\d/i.test(String(r[k] || "").trim())),
      );
    }

    const all = rows
      .map((r) => ({
        name: String(r[nameKey] || "").trim(),
        organization: orgKey ? String(r[orgKey] || "").trim() : "",
        round: roundKey ? String(r[roundKey] || "").trim() : "",
      }))
      .filter((r) => r.name);

    let filtered = all;
    if (ROUND_FILTER && roundKey) {
      const want = ROUND_FILTER.toLowerCase();
      filtered = all.filter((r) => r.round.toLowerCase() === want);
    }

    return {
      rows: filtered.map(({ name, organization }) => ({ name, organization })),
      totalBeforeFilter: all.length,
    };
  };

  // Try loading the responses spreadsheet via fetch (works when served over http/https)
  const SOURCE_FILE = "Boston Gates (Responses).xlsx";
  const tryLoadSourceXlsx = async () => {
    try {
      const res = await fetch(encodeURI(SOURCE_FILE), { cache: "no-store" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      const buf = await res.arrayBuffer();
      const { rows, totalBeforeFilter } = parseSpreadsheet(buf);
      if (rows.length) {
        setAttendees(rows, SOURCE_FILE, totalBeforeFilter);
        return true;
      }
    } catch (e) {
      // Likely opened via file:// — fall back to bundled list
    }
    return false;
  };

  // ----- Spinner / shuffle animation -----
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  const pickRandom = () =>
    attendees[Math.floor(Math.random() * attendees.length)];

  const showOnReel = (entry) => {
    reelName.textContent = entry.name;
    reelOrg.textContent = entry.organization || "\u00A0";
  };

  const setReelClass = (cls) => {
    reel.classList.remove("spinning", "fast", "medium", "slow", "reveal");
    if (cls) reel.classList.add(cls);
  };

  const spinAndPick = async () => {
    if (isSpinning || attendees.length === 0) return;
    isSpinning = true;
    pickBtn.disabled = true;
    resetBtn.hidden = true;
    overlay.hidden = true;

    const winner = pickRandom();

    // Phase 1: very fast shuffle
    reel.classList.add("spinning", "fast");
    let elapsed = 0;
    let interval = 55;
    const fastDuration = 1800;
    while (elapsed < fastDuration) {
      showOnReel(pickRandom());
      await sleep(interval);
      elapsed += interval;
    }

    // Phase 2: medium speed
    setReelClass("spinning");
    reel.classList.add("medium");
    elapsed = 0;
    const mediumDuration = 1400;
    interval = 110;
    while (elapsed < mediumDuration) {
      showOnReel(pickRandom());
      await sleep(interval);
      elapsed += interval;
      interval += 12;
    }

    // Phase 3: slow countdown to the winner
    setReelClass("spinning");
    reel.classList.add("slow");
    const slowSteps = 6;
    for (let i = 0; i < slowSteps; i++) {
      showOnReel(pickRandom());
      await sleep(180 + i * 80);
    }

    // Final reveal
    setReelClass("reveal");
    showOnReel(winner);
    burstConfetti();

    await sleep(900);
    showWinner(winner);

    isSpinning = false;
    resetBtn.hidden = false;
    pickBtn.disabled = false;
  };

  // ----- Confetti -----
  const burstConfetti = () => {
    if (typeof confetti !== "function") return;
    // ISTE 2026 brand palette
    const colors = ["#5753FA", "#E83FA4", "#70AD47", "#F44E2B", "#CCE9FE", "#1D2B63"];
    confetti({
      particleCount: 140,
      spread: 80,
      origin: { y: 0.45 },
      colors,
      scalar: 1.05,
    });
    setTimeout(() => {
      confetti({ particleCount: 70, angle: 60, spread: 70, origin: { x: 0, y: 0.6 }, colors });
      confetti({ particleCount: 70, angle: 120, spread: 70, origin: { x: 1, y: 0.6 }, colors });
    }, 220);
  };

  const sustainedConfetti = (durationMs = 2500) => {
    if (typeof confetti !== "function") return;
    const end = Date.now() + durationMs;
    const colors = ["#5753FA", "#E83FA4", "#70AD47", "#F44E2B", "#CCE9FE"];
    (function frame() {
      confetti({
        particleCount: 4,
        angle: 60,
        spread: 65,
        startVelocity: 55,
        origin: { x: 0, y: 0.7 },
        colors,
      });
      confetti({
        particleCount: 4,
        angle: 120,
        spread: 65,
        startVelocity: 55,
        origin: { x: 1, y: 0.7 },
        colors,
      });
      if (Date.now() < end) requestAnimationFrame(frame);
    })();
  };

  // ----- Winner overlay -----
  const showWinner = (entry) => {
    winnerName.textContent = entry.name;
    winnerOrg.textContent = entry.organization || "";
    overlay.hidden = false;
    sustainedConfetti(2200);
  };

  const closeOverlay = () => {
    overlay.hidden = true;
  };

  // ----- Event bindings -----
  pickBtn.addEventListener("click", spinAndPick);
  resetBtn.addEventListener("click", () => {
    setReelClass("");
    reelName.textContent = "Press the button to start!";
    reelOrg.innerHTML = "&nbsp;";
    resetBtn.hidden = true;
    overlay.hidden = true;
  });
  closeWinner.addEventListener("click", closeOverlay);
  spinAgain.addEventListener("click", () => {
    closeOverlay();
    spinAndPick();
  });
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) closeOverlay();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeOverlay();
    if (e.key === " " || e.key === "Enter") {
      const tag = (document.activeElement && document.activeElement.tagName) || "";
      if (tag === "BUTTON" || tag === "INPUT") return;
      e.preventDefault();
      spinAndPick();
    }
  });

  fileInput.addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const { rows, totalBeforeFilter } = parseSpreadsheet(buf);
      if (!rows.length) {
        const tag = ROUND_FILTER ? ` tagged "${ROUND_FILTER}"` : "";
        setHint(
          `Couldn't find any attendees${tag} in <code>${file.name}</code>. Make sure it has a Name column${ROUND_FILTER ? " and a round/group column" : ""}.`,
          true,
        );
        return;
      }
      setAttendees(rows, file.name, totalBeforeFilter);
      setReelClass("");
      reelName.textContent = "Press the button to start!";
      reelOrg.innerHTML = "&nbsp;";
    } catch (err) {
      console.error(err);
      setHint(`Something went wrong reading <code>${file.name}</code>.`, true);
    }
  });

  // ----- Boot -----
  pickBtn.disabled = true;
  (async () => {
    const ok = await tryLoadSourceXlsx();
    if (!ok) {
      setAttendees(
        FALLBACK_ATTENDEES,
        SOURCE_FILE + " (bundled)",
        FALLBACK_ATTENDEES.length,
      );
    }
  })();
})();
