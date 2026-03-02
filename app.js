const STORAGE_KEY = "school-scheduler-v3";
    const state = {
      rooms: [],
      slots: [],
      teachers: [],
      loads: {},
      teacherBlocks: {},
      lockedLessons: {},
      lessonInstances: [],
      schedule: null,
      nextId: 1
    };

    const dayOrder = { Segunda: 1, "Terça": 2, Quarta: 3, Quinta: 4, Sexta: 5, "Sábado": 6 };

    const roomForm = document.getElementById("room-form");
    const slotForm = document.getElementById("slot-form");
    const slotBatchForm = document.getElementById("slot-batch-form");
    const teacherForm = document.getElementById("teacher-form");
    const loadForm = document.getElementById("load-form");
    const blockForm = document.getElementById("block-form");
    const buildForm = document.getElementById("build-form");
    const clearAllBtn = document.getElementById("clear-all-btn");
    const regenerateBtn = document.getElementById("regenerate-btn");

    const roomList = document.getElementById("room-list");
    const removeRoomBtn = document.getElementById("remove-room-btn");
    const slotSelect = document.getElementById("slot-select");
    const removeSlotBtn = document.getElementById("remove-slot-btn");
    const teacherList = document.getElementById("teacher-list");
    const removeTeacherBtn = document.getElementById("remove-teacher-btn");
    const slotRoom = document.getElementById("slot-room");
    const loadList = document.getElementById("load-list");
    const blockList = document.getElementById("block-list");
    const alertList = document.getElementById("alert-list");

    const loadTeacher = document.getElementById("load-teacher");
    const loadRoom = document.getElementById("load-room");
    const loadQty = document.getElementById("load-qty");

    const blockTeacher = document.getElementById("block-teacher");
    const blockDay = document.getElementById("block-day");
    const blockStart = document.getElementById("block-start");
    const blockEnd = document.getElementById("block-end");
    const blockFullDay = document.getElementById("block-full-day");

    const viewMode = document.getElementById("view-mode");
    const roomViewWrap = document.getElementById("room-view-wrap");
    const viewRoom = document.getElementById("view-room");
    const exportCsvBtn = document.getElementById("export-csv-btn");
    const exportPdfBtn = document.getElementById("export-pdf-btn");

    const scheduleOutput = document.getElementById("schedule-output");
    const statusEl = document.getElementById("status");
    const generationProgress = document.getElementById("generation-progress");
    const progressTitle = document.getElementById("progress-title");
    const progressPercent = document.getElementById("progress-percent");
    const progressBar = document.getElementById("progress-bar");
    const progressMeta = document.getElementById("progress-meta");
    let runtimeAlertMessage = "";
    let schedulerWorker = null;
    let scheduleRequestCounter = 0;
    let pendingScheduleRequest = null;
    let progressIntervalId = null;
    let progressStartedAt = 0;
    let progressValue = 0;
    let progressStage = "Preparando";
    let progressLastUpdateAt = 0;

    function makeId(prefix) { return `${prefix}-${state.nextId++}`; }
    function keyLoad(teacherId, roomId) { return `${teacherId}__${roomId}`; }
    function slotLabel(slot) { return `${slot.day} ${slot.start}-${slot.end}`; }
    function timeToMinutes(value) { const [h, m] = value.split(":").map(Number); return h * 60 + m; }
    function minutesToTime(total) { return `${String(Math.floor(total / 60)).padStart(2, "0")}:${String(total % 60).padStart(2, "0")}`; }
    function slotTimeKey(slot) { return `${slot.day}@@${slot.start}@@${slot.end}`; }
    function roomSlots(roomId) { return state.slots.filter((s) => (s.roomId || "") === roomId); }

    function setStatus(message, type) {
      statusEl.className = `status ${type}`;
      statusEl.textContent = message;
    }

    function clearStatus() {
      statusEl.className = "status";
      statusEl.textContent = "";
    }

    function renderProgress() {
      if (!generationProgress || generationProgress.hidden) return;
      const elapsedSec = Math.max(0, Math.floor((Date.now() - progressStartedAt) / 1000));
      progressPercent.textContent = `${Math.max(0, Math.min(100, Math.round(progressValue)))}%`;
      progressBar.style.width = `${Math.max(0, Math.min(100, progressValue))}%`;
      progressMeta.textContent = `${progressStage} | ${elapsedSec}s`;
    }

    function showGenerationProgress(title = "Processando geração") {
      if (!generationProgress) return;
      progressStartedAt = Date.now();
      progressValue = 2;
      progressStage = "Preparando dados";
      progressLastUpdateAt = Date.now();
      progressTitle.textContent = title;
      generationProgress.hidden = false;
      if (progressIntervalId) clearInterval(progressIntervalId);
      progressIntervalId = setInterval(() => {
        const idleMs = Date.now() - progressLastUpdateAt;
        if (idleMs > 450 && progressValue < 88) {
          progressValue = Math.min(88, progressValue + 1);
        }
        if (idleMs > 3000 && progressValue >= 70) {
          progressStage = "Processamento pesado";
        }
        renderProgress();
      }, 250);
      renderProgress();
    }

    function updateGenerationProgress(percent, stage) {
      if (!generationProgress || generationProgress.hidden) return;
      if (Number.isFinite(percent)) {
        const safe = Math.max(2, Math.min(99, percent));
        progressValue = Math.max(progressValue, safe);
      }
      if (stage) progressStage = stage;
      progressLastUpdateAt = Date.now();
      renderProgress();
    }

    function hideGenerationProgress() {
      if (!generationProgress) return;
      if (progressIntervalId) {
        clearInterval(progressIntervalId);
        progressIntervalId = null;
      }
      generationProgress.hidden = true;
      progressValue = 0;
      progressStage = "Preparando";
      progressBar.style.width = "0%";
      progressPercent.textContent = "0%";
      progressMeta.textContent = "Preparando...";
    }

    function ensureSchedulerWorker() {
      if (!("Worker" in window)) return null;
      if (schedulerWorker) return schedulerWorker;
      schedulerWorker = new Worker("scheduler-worker.js?v=20260302d");
      schedulerWorker.addEventListener("message", (event) => {
        const data = event.data || {};
        if (!pendingScheduleRequest) return;
        if (data.requestId !== pendingScheduleRequest.id) return;

        if (data.type === "progress") {
          const current = pendingScheduleRequest;
          if (typeof current.onProgress === "function") current.onProgress(data.progress || {});
          return;
        }

        const current = pendingScheduleRequest;
        pendingScheduleRequest = null;
        if (current.timeoutId) clearTimeout(current.timeoutId);
        if (data.type === "result") {
          current.resolve(data.result);
          return;
        }
        current.reject(new Error(data.error || "Falha no worker de geração."));
      });
      schedulerWorker.addEventListener("error", (event) => {
        if (!pendingScheduleRequest) return;
        const current = pendingScheduleRequest;
        pendingScheduleRequest = null;
        if (current.timeoutId) clearTimeout(current.timeoutId);
        current.reject(new Error(event.message || "Erro inesperado no worker de geração."));
      });
      return schedulerWorker;
    }

    function resetSchedulerWorker() {
      if (schedulerWorker) schedulerWorker.terminate();
      schedulerWorker = null;
    }

    function runGeneration(tasks, fixedAssignments = new Map(), onProgress) {
      const worker = ensureSchedulerWorker();
      if (!worker) {
        if (typeof onProgress === "function") onProgress({ percent: 25, stage: "Processando localmente" });
        const local = generateWithPreferences(tasks, fixedAssignments);
        if (typeof onProgress === "function") onProgress({ percent: 95, stage: "Finalizando" });
        return Promise.resolve(local);
      }

      if (pendingScheduleRequest) {
        pendingScheduleRequest.reject(new Error("generation-cancelled"));
        pendingScheduleRequest = null;
        resetSchedulerWorker();
      }

      const activeWorker = ensureSchedulerWorker();
      const requestId = ++scheduleRequestCounter;
      const payload = {
        type: "generate",
        requestId,
        tasks,
        slots: state.slots,
        teachers: state.teachers,
        teacherBlocks: state.teacherBlocks,
        fixedAssignments: Array.from(fixedAssignments.entries())
      };
      const complexity = tasks.length + state.slots.length + state.teachers.length;
      const hardTimeoutMs = Math.max(6000, Math.min(18000, 5000 + complexity * 12));

      return new Promise((resolve, reject) => {
        const timeoutId = setTimeout(() => {
          if (!pendingScheduleRequest || pendingScheduleRequest.id !== requestId) return;
          pendingScheduleRequest = null;
          resetSchedulerWorker();
          reject(new Error("generation-timeout"));
        }, hardTimeoutMs);

        pendingScheduleRequest = { id: requestId, resolve, reject, onProgress, timeoutId };
        try {
          activeWorker.postMessage(payload);
        } catch (err) {
          pendingScheduleRequest = null;
          clearTimeout(timeoutId);
          resetSchedulerWorker();
          reject(err);
        }
      });
    }

    function setGenerationBusy(isBusy) {
      const buildBtn = buildForm.querySelector("button[type='submit']");
      if (buildBtn) buildBtn.disabled = isBusy;
      regenerateBtn.disabled = isBusy || !state.lessonInstances.length;
    }

    function roomParts(name) {
      return { number: Number(name[0]), letter: name[1] };
    }

    function sortRooms() {
      state.rooms.sort((a, b) => {
        const pa = roomParts(a.name);
        const pb = roomParts(b.name);
        if (pa.number !== pb.number) return pa.number - pb.number;
        return pa.letter.localeCompare(pb.letter);
      });
    }

    function sortSlots() {
      state.slots.sort((a, b) => {
        if (dayOrder[a.day] !== dayOrder[b.day]) return dayOrder[a.day] - dayOrder[b.day];
        if (a.start !== b.start) return a.start.localeCompare(b.start);
        if (a.end !== b.end) return a.end.localeCompare(b.end);
        const roomA = state.rooms.find((r) => r.id === a.roomId)?.name || "";
        const roomB = state.rooms.find((r) => r.id === b.roomId)?.name || "";
        return roomA.localeCompare(roomB);
      });
    }

    function sortTeachers() {
      state.teachers.sort((a, b) => a.name.localeCompare(b.name, "pt-BR", { sensitivity: "base" }));
    }

    function saveState() {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify({
          state,
          ui: { viewMode: viewMode.value, viewRoom: viewRoom.value }
        }));
      } catch (_e) {
      }
    }

    function loadState() {
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return;
        const parsed = JSON.parse(raw);
        const saved = parsed.state || {};

        state.rooms = Array.isArray(saved.rooms) ? saved.rooms : [];
        state.slots = Array.isArray(saved.slots) ? saved.slots : [];
        state.teachers = Array.isArray(saved.teachers) ? saved.teachers : [];
        state.loads = saved.loads && typeof saved.loads === "object" ? saved.loads : {};
        state.teacherBlocks = saved.teacherBlocks && typeof saved.teacherBlocks === "object" ? saved.teacherBlocks : {};
        state.lockedLessons = saved.lockedLessons && typeof saved.lockedLessons === "object" ? saved.lockedLessons : {};
        state.lessonInstances = Array.isArray(saved.lessonInstances) ? saved.lessonInstances : [];
        state.schedule = saved.schedule && typeof saved.schedule === "object" ? saved.schedule : null;
        state.nextId = Number.isInteger(saved.nextId) && saved.nextId > 0 ? saved.nextId : 1;

        sortRooms();
        sortTeachers();
        if (state.rooms.length) {
          for (const slot of state.slots) {
            if (!slot.roomId) slot.roomId = state.rooms[0].id;
          }
        }
        sortSlots();

        if (parsed.ui?.viewMode === "room" || parsed.ui?.viewMode === "global") viewMode.value = parsed.ui.viewMode;
      } catch (_e) {
      }
    }

    function removeEntity(list, id) {
      const idx = list.findIndex((x) => x.id === id);
      if (idx >= 0) list.splice(idx, 1);
      state.lessonInstances = [];
      state.schedule = null;
      clearStatus();
    }

    function normalizeLoads() {
      for (const key of Object.keys(state.loads)) {
        const [teacherId, roomId] = key.split("__");
        const hasTeacher = state.teachers.some((t) => t.id === teacherId);
        const hasRoom = state.rooms.some((r) => r.id === roomId);
        if (!hasTeacher || !hasRoom || state.loads[key] <= 0) delete state.loads[key];
      }
    }

    function normalizeBlocks() {
      for (const teacherId of Object.keys(state.teacherBlocks)) {
        if (!state.teachers.some((t) => t.id === teacherId)) {
          delete state.teacherBlocks[teacherId];
          continue;
        }
        const blocks = Array.isArray(state.teacherBlocks[teacherId]) ? state.teacherBlocks[teacherId] : [];
        state.teacherBlocks[teacherId] = blocks.filter((b) => {
          if (!dayOrder[b.day]) return false;
          if (b.fullDay) return true;
          return b.end > b.start;
        });
      }
    }

    function normalizeLessonInstances() {
      const validTeachers = new Set(state.teachers.map((t) => t.id));
      const validRooms = new Set(state.rooms.map((r) => r.id));
      const validSlots = new Set(state.slots.map((s) => s.id));

      state.lessonInstances = (state.lessonInstances || []).filter((lesson) => {
        if (!validTeachers.has(lesson.teacherId)) return false;
        if (!validRooms.has(lesson.roomId)) return false;
        if (lesson.slotId && !validSlots.has(lesson.slotId)) lesson.slotId = null;
        return true;
      });
    }

    function isLessonLocked(lessonId) {
      return Boolean(state.lockedLessons?.[lessonId]);
    }

    function normalizeLockedLessons() {
      const valid = new Set(state.lessonInstances.filter((l) => l.slotId).map((l) => l.lessonId));
      for (const lessonId of Object.keys(state.lockedLessons || {})) {
        if (!valid.has(lessonId)) delete state.lockedLessons[lessonId];
      }
    }

    function detectConflicts() {
      const conflicts = [];
      const teacherTimeMap = new Map();
      const roomSlotMap = new Map();
      const slotById = new Map(state.slots.map((slot) => [slot.id, slot]));

      for (const lesson of state.lessonInstances) {
        if (!lesson.slotId) continue;
        const slot = slotById.get(lesson.slotId);
        if (!slot) continue;

        const tKey = `${lesson.teacherId}@@${slot.day}@@${slot.start}@@${slot.end}`;
        if (!teacherTimeMap.has(tKey)) teacherTimeMap.set(tKey, []);
        teacherTimeMap.get(tKey).push({ lesson, slot });

        const rKey = `${lesson.roomId}@@${slot.id}`;
        if (!roomSlotMap.has(rKey)) roomSlotMap.set(rKey, []);
        roomSlotMap.get(rKey).push({ lesson, slot });

        if (isTeacherBlocked(lesson.teacherId, slot)) {
          conflicts.push(`Professor ${lesson.teacherName} alocado em horário bloqueado: ${slot.day} ${slot.start}-${slot.end} (Sala ${lesson.roomName}).`);
        }
      }

      for (const entries of teacherTimeMap.values()) {
        if (entries.length <= 1) continue;
        const { lesson, slot } = entries[0];
        const rooms = entries.map((e) => e.lesson.roomName).join(", ");
        conflicts.push(`Professor ${lesson.teacherName} com conflito em ${slot.day} ${slot.start}-${slot.end}: salas ${rooms}.`);
      }

      for (const entries of roomSlotMap.values()) {
        if (entries.length <= 1) continue;
        const { lesson, slot } = entries[0];
        const teachers = entries.map((e) => e.lesson.teacherName).join(", ");
        conflicts.push(`Sala ${lesson.roomName} com mais de um professor no mesmo horário (${slot.day} ${slot.start}-${slot.end}): ${teachers}.`);
      }

      return conflicts;
    }

    function renderAlerts() {
      const conflicts = detectConflicts();
      alertList.innerHTML = "";
      if (!conflicts.length) {
        const li = document.createElement("li");
        li.textContent = "Nenhum conflito identificado no momento.";
        alertList.appendChild(li);
        return;
      }
      const uniqueConflicts = [...new Set(conflicts)];
      for (const msg of uniqueConflicts) {
        const li = document.createElement("li");
        li.textContent = msg;
        alertList.appendChild(li);
      }
    }

    function normalizeRoomData() {
      const validRooms = new Set(state.rooms.map((r) => r.id));
      state.slots = state.slots.filter((slot) => validRooms.has(slot.roomId));
      sortSlots();
    }

    function rebuildScheduleFromLessons() {
      const schedule = {};
      const slotById = new Map(state.slots.map((slot) => [slot.id, slot]));
      for (const room of state.rooms) schedule[room.id] = {};

      for (const lesson of state.lessonInstances) {
        if (!lesson.slotId) continue;
        const slot = slotById.get(lesson.slotId);
        if (!slot) continue;
        schedule[lesson.roomId][lesson.slotId] = {
          teacherId: lesson.teacherId,
          teacherName: lesson.teacherName,
          day: slot.day,
          start: slot.start,
          end: slot.end
        };
      }

      state.schedule = schedule;
    }

    function addSlot(roomId, day, start, end) {
      if (!roomId) return { ok: false, reason: "Selecione uma série/sala." };
      if (!start || !end) return { ok: false, reason: "Informe horário de início e fim." };
      if (end <= start) return { ok: false, reason: "O horário de fim deve ser maior que o de início." };
      const exists = state.slots.some((s) => s.roomId === roomId && s.day === day && s.start === start && s.end === end);
      if (exists) return { ok: false, reason: "duplicate" };
      state.slots.push({ id: makeId("slot"), roomId, day, start, end });
      return { ok: true };
    }

    function groupedSlotsByDay() {
      const grouped = [];
      let current = null;
      for (const slot of state.slots) {
        if (!current || current.day !== slot.day) {
          current = { day: slot.day, slots: [] };
          grouped.push(current);
        }
        current.slots.push(slot);
      }
      return grouped;
    }

    function renderSimpleList(el, items, textFn, removeFn, emptyText) {
      el.innerHTML = "";
      if (!items.length) {
        const li = document.createElement("li");
        li.textContent = emptyText;
        el.appendChild(li);
        return;
      }

      items.forEach((item) => {
        const li = document.createElement("li");
        const span = document.createElement("span");
        span.textContent = textFn(item);
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "remove-btn";
        btn.textContent = "Remover";
        btn.addEventListener("click", () => removeFn(item.id));
        li.appendChild(span);
        li.appendChild(btn);
        el.appendChild(li);
      });
    }

    function renderRooms() {
      const prevRoom = roomList.value;
      const options = state.rooms
        .map((room) => `<option value="${room.id}">${room.name}</option>`)
        .join("");
      roomList.innerHTML = options || "<option value=''>Nenhuma sala cadastrada</option>";
      roomList.disabled = !state.rooms.length;
      removeRoomBtn.disabled = !state.rooms.length;
      if (state.rooms.length) {
        const selectedRoomId = prevRoom || roomList.value || state.rooms[0].id;
        roomList.value = selectedRoomId;
      }
    }

    function renderSelectors() {
      const prevLoadTeacher = loadTeacher.value;
      const prevBlockTeacher = blockTeacher.value;
      const prevLoadRoom = loadRoom.value;
      const prevSlotRoom = slotRoom.value;
      const prevTeacherList = teacherList.value;
      const prevViewRoom = viewRoom.value;
      const teachers = state.teachers.map((t) => `<option value="${t.id}">${t.name}</option>`).join("");
      const rooms = state.rooms.map((r) => `<option value="${r.id}">${r.name}</option>`).join("");

      loadTeacher.innerHTML = teachers || "<option value=''>Sem professores</option>";
      blockTeacher.innerHTML = teachers || "<option value=''>Sem professores</option>";
      loadRoom.innerHTML = rooms || "<option value=''>Sem salas</option>";
      slotRoom.innerHTML = rooms || "<option value=''>Sem salas</option>";
      teacherList.innerHTML = teachers || "<option value=''>Nenhum professor cadastrado</option>";

      viewRoom.innerHTML = rooms;
      if (prevLoadTeacher) loadTeacher.value = prevLoadTeacher;
      if (prevBlockTeacher) blockTeacher.value = prevBlockTeacher;
      if (prevLoadRoom) loadRoom.value = prevLoadRoom;
      if (prevSlotRoom) slotRoom.value = prevSlotRoom;
      if (prevTeacherList) teacherList.value = prevTeacherList;
      if (prevViewRoom) viewRoom.value = prevViewRoom;
      if (!slotRoom.value && state.rooms[0]) slotRoom.value = state.rooms[0].id;
      roomViewWrap.style.display = viewMode.value === "room" ? "grid" : "none";

      const loadDisabled = !state.teachers.length || !state.rooms.length;
      loadTeacher.disabled = loadDisabled;
      loadRoom.disabled = loadDisabled;
      loadQty.disabled = loadDisabled;
      loadForm.querySelector("button").disabled = loadDisabled;

      const blockDisabled = !state.teachers.length;
      blockTeacher.disabled = blockDisabled;
      blockDay.disabled = blockDisabled;
      blockStart.disabled = blockDisabled || blockFullDay.checked;
      blockEnd.disabled = blockDisabled || blockFullDay.checked;
      blockForm.querySelector("button").disabled = blockDisabled;

      const slotDisabled = !state.rooms.length;
      slotRoom.disabled = slotDisabled;
      slotForm.querySelector("button").disabled = slotDisabled;
      slotBatchForm.querySelector("button").disabled = slotDisabled;

      teacherList.disabled = !state.teachers.length;
      removeTeacherBtn.disabled = !state.teachers.length;
    }

    function renderSlotSelector() {
      const prevSlot = slotSelect.value;
      const roomById = new Map(state.rooms.map((room) => [room.id, room.name]));
      const options = state.slots
        .map((s) => {
          const roomName = roomById.get(s.roomId) || "Sala";
          return `<option value="${s.id}">${roomName} | ${slotLabel(s)}</option>`;
        })
        .join("");
      slotSelect.innerHTML = options || "<option value=''>Nenhum horário cadastrado</option>";
      if (prevSlot) slotSelect.value = prevSlot;
      const disabled = !state.slots.length;
      slotSelect.disabled = disabled;
      removeSlotBtn.disabled = disabled;
    }

    function renderLoads() {
      normalizeLoads();
      const teacherById = new Map(state.teachers.map((teacher) => [teacher.id, teacher]));
      const roomById = new Map(state.rooms.map((room) => [room.id, room]));
      const entries = Object.entries(state.loads)
        .map(([key, qty]) => {
          const [teacherId, roomId] = key.split("__");
          const teacher = teacherById.get(teacherId);
          const room = roomById.get(roomId);
          if (!teacher || !room) return null;
          return { key, teacher, room, qty };
        })
        .filter(Boolean)
        .sort((a, b) => a.teacher.name.localeCompare(b.teacher.name) || a.room.name.localeCompare(b.room.name));

      loadList.innerHTML = "";
      if (!entries.length) {
        const li = document.createElement("li");
        li.textContent = "Nenhuma carga cadastrada.";
        loadList.appendChild(li);
        return;
      }

      const byTeacher = new Map();
      for (const e of entries) {
        if (!byTeacher.has(e.teacher.id)) byTeacher.set(e.teacher.id, { teacher: e.teacher, items: [] });
        byTeacher.get(e.teacher.id).items.push(e);
      }

      for (const group of byTeacher.values()) {
        const li = document.createElement("li");
        li.style.display = "block";
        const details = document.createElement("details");
        details.className = "teacher-load-details";

        const summary = document.createElement("summary");
        const total = group.items.reduce((acc, item) => acc + item.qty, 0);
        summary.textContent = `${group.teacher.name} (${total} aula(s))`;
        details.appendChild(summary);

        const inner = document.createElement("div");
        inner.className = "teacher-load-items";
        for (const item of group.items) {
          const row = document.createElement("div");
          row.className = "series-item";
          const text = document.createElement("span");
          text.textContent = `${item.room.name}: ${item.qty} aula(s)`;
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "remove-btn";
          btn.textContent = "Remover";
          btn.addEventListener("click", () => {
            delete state.loads[item.key];
            state.lessonInstances = [];
            state.schedule = null;
            clearStatus();
            renderAll();
          });
          row.appendChild(text);
          row.appendChild(btn);
          inner.appendChild(row);
        }
        details.appendChild(inner);
        li.appendChild(details);
        loadList.appendChild(li);
      }
    }

    function renderBlocks() {
      normalizeBlocks();
      const entries = [];
      for (const teacher of state.teachers) {
        for (const b of (state.teacherBlocks[teacher.id] || [])) {
          entries.push({ teacher, block: b });
        }
      }

      entries.sort((a, b) =>
        a.teacher.name.localeCompare(b.teacher.name) ||
        dayOrder[a.block.day] - dayOrder[b.block.day] ||
        a.block.start.localeCompare(b.block.start)
      );

      blockList.innerHTML = "";
      if (!entries.length) {
        const li = document.createElement("li");
        li.textContent = "Nenhum bloqueio cadastrado.";
        blockList.appendChild(li);
        return;
      }

      for (const e of entries) {
        const li = document.createElement("li");
        const span = document.createElement("span");
        span.textContent = e.block.fullDay
          ? `${e.teacher.name} | ${e.block.day} (dia inteiro)`
          : `${e.teacher.name} | ${e.block.day} ${e.block.start}-${e.block.end}`;
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "remove-btn";
        btn.textContent = "Remover";
        btn.addEventListener("click", () => {
          state.teacherBlocks[e.teacher.id] = (state.teacherBlocks[e.teacher.id] || []).filter((x) => x.id !== e.block.id);
          state.lessonInstances = [];
          state.schedule = null;
          clearStatus();
          renderAll();
        });
        li.appendChild(span);
        li.appendChild(btn);
        blockList.appendChild(li);
      }
    }

    function buildLessonsFromLoads() {
      const lessons = [];
      for (const teacher of state.teachers) {
        for (const room of state.rooms) {
          const qty = state.loads[keyLoad(teacher.id, room.id)] || 0;
          for (let i = 0; i < qty; i++) {
            lessons.push({
              lessonId: makeId("lesson"),
              teacherId: teacher.id,
              teacherName: teacher.name,
              roomId: room.id,
              roomName: room.name,
              slotId: null
            });
          }
        }
      }
      return lessons;
    }

    function isTeacherBlocked(teacherId, slot) {
      const blocks = state.teacherBlocks[teacherId] || [];
      const slotStart = timeToMinutes(slot.start);
      const slotEnd = timeToMinutes(slot.end);
      return blocks.some((b) => {
        if (b.day !== slot.day) return false;
        if (b.fullDay) return true;
        const bStart = timeToMinutes(b.start);
        const bEnd = timeToMinutes(b.end);
        return slotStart < bEnd && bStart < slotEnd;
      });
    }

    function canAllocate(tasks) {
      if (!state.rooms.length) return "Cadastre ao menos 1 sala.";
      if (!state.teachers.length) return "Cadastre ao menos 1 professor.";
      if (!state.slots.length) return "Cadastre ao menos 1 horário.";
      if (!tasks.length) return "Defina ao menos uma carga de aula.";

      const totalCapacity = state.slots.length;
      const byTeacher = new Map();
      const byRoom = new Map();
      for (const t of tasks) {
        byTeacher.set(t.teacherId, (byTeacher.get(t.teacherId) || 0) + 1);
        byRoom.set(t.roomId, (byRoom.get(t.roomId) || 0) + 1);
      }

      const teacherIssues = [];
      const roomIssues = [];
      const teacherById = new Map(state.teachers.map((teacher) => [teacher.id, teacher.name]));
      const roomById = new Map(state.rooms.map((room) => [room.id, room.name]));
      const roomSlotCount = new Map();
      const blockedKeyByTeacher = new Map();

      for (const slot of state.slots) {
        roomSlotCount.set(slot.roomId, (roomSlotCount.get(slot.roomId) || 0) + 1);
      }

      for (const [teacherId, count] of byTeacher) {
        if (!blockedKeyByTeacher.has(teacherId)) {
          const keys = new Set();
          for (const slot of state.slots) {
            if (!isTeacherBlocked(teacherId, slot)) keys.add(slotTimeKey(slot));
          }
          blockedKeyByTeacher.set(teacherId, keys);
        }
        const available = blockedKeyByTeacher.get(teacherId).size;
        if (count > available) {
          const name = teacherById.get(teacherId) || teacherId;
          teacherIssues.push(`${name}: ${count} aula(s) para ${available} horário(s) disponível(is)`);
        }
      }

      for (const [roomId, count] of byRoom) {
        const availableSlots = roomSlotCount.get(roomId) || 0;
        if (count > availableSlots) {
          const room = roomById.get(roomId) || roomId;
          roomIssues.push(`Sala ${room}: ${count} aula(s) para ${availableSlots} horário(s) cadastrado(s)`);
        }
      }

      if (tasks.length > totalCapacity || teacherIssues.length || roomIssues.length) {
        const parts = [`Aulas totais ${tasks.length} / capacidade total ${totalCapacity}.`];
        if (teacherIssues.length) parts.push(`Professores excedendo disponibilidade: ${teacherIssues.join("; ")}.`);
        if (roomIssues.length) parts.push(`Salas sobrecarregadas: ${roomIssues.join("; ")}.`);
        return parts.join(" ");
      }

      return null;
    }

    function buildDayTimeIndexMap(slotById) {
      const dayTimes = new Map();
      for (const slot of slotById.values()) {
        if (!dayTimes.has(slot.day)) dayTimes.set(slot.day, new Set());
        dayTimes.get(slot.day).add(`${slot.start}-${slot.end}`);
      }
      const indexMap = new Map();
      for (const [day, times] of dayTimes.entries()) {
        const ordered = Array.from(times).sort((a, b) => a.split("-")[0].localeCompare(b.split("-")[0]));
        const map = new Map();
        for (let i = 0; i < ordered.length; i++) map.set(ordered[i], i);
        indexMap.set(day, map);
      }
      return indexMap;
    }

    function cloneTeacherStateMap(source) {
      const out = new Map();
      for (const [teacherId, state] of source.entries()) {
        const dayIndices = new Map();
        for (const [day, indices] of state.dayIndices.entries()) dayIndices.set(day, new Set(indices));
        const dayCounts = new Map(state.dayCounts.entries());
        const daySpan = new Map();
        for (const [day, span] of state.daySpan.entries()) daySpan.set(day, { min: span.min, max: span.max, size: span.size });
        out.set(teacherId, { dayIndices, dayCounts, daySpan });
      }
      return out;
    }

    function generateSchedule(tasks, options = {}) {
      generateSchedule.lastError = "";
      const slotById = new Map(state.slots.map((s) => [s.id, s]));
      const dayTimeIndexMap = buildDayTimeIndexMap(slotById);
      const fixedAssignments = options.fixedAssignments instanceof Map ? options.fixedAssignments : new Map();
      const maxConsecutive = Number.isInteger(options.maxConsecutive) && options.maxConsecutive > 0 ? options.maxConsecutive : null;
      const avoidGaps = Boolean(options.avoidGaps);
      const balanceDaily = Boolean(options.balanceDaily);
      const deadline = Date.now() + (Number.isFinite(options.timeLimitMs) ? options.timeLimitMs : 1800);
      const maxSteps = Number.isFinite(options.maxSteps) ? options.maxSteps : 350000;
      const slotKeyById = new Map();
      const roomSlotIds = new Map();
      const blockedSlotsByTeacher = new Map();

      for (const slot of state.slots) {
        slotKeyById.set(slot.id, slotTimeKey(slot));
        if (!roomSlotIds.has(slot.roomId)) roomSlotIds.set(slot.roomId, []);
        roomSlotIds.get(slot.roomId).push(slot.id);
      }
      for (const teacher of state.teachers) {
        const blocked = new Set();
        for (const slot of state.slots) {
          if (isTeacherBlocked(teacher.id, slot)) blocked.add(slot.id);
        }
        blockedSlotsByTeacher.set(teacher.id, blocked);
      }

      const usedTeacherSlot = new Set();
      const usedRoomSlot = new Set();
      const assignment = new Map();
      const teacherState = new Map();
      const taskCandidates = new Map();
      let steps = 0;
      let aborted = false;

      function getTeacherState(teacherId) {
        if (!teacherState.has(teacherId)) teacherState.set(teacherId, { dayIndices: new Map(), dayCounts: new Map(), daySpan: new Map() });
        return teacherState.get(teacherId);
      }

      function addTeacherSlot(teacherId, slot) {
        const st = getTeacherState(teacherId);
        const timeKey = `${slot.start}-${slot.end}`;
        const idx = dayTimeIndexMap.get(slot.day)?.get(timeKey);
        if (Number.isInteger(idx)) {
          if (!st.dayIndices.has(slot.day)) st.dayIndices.set(slot.day, new Set());
          const daySet = st.dayIndices.get(slot.day);
          daySet.add(idx);
          const span = st.daySpan.get(slot.day);
          if (!span) {
            st.daySpan.set(slot.day, { min: idx, max: idx, size: 1 });
          } else {
            span.min = Math.min(span.min, idx);
            span.max = Math.max(span.max, idx);
            span.size += 1;
          }
        }
        st.dayCounts.set(slot.day, (st.dayCounts.get(slot.day) || 0) + 1);
      }

      function removeTeacherSlot(teacherId, slot) {
        const st = getTeacherState(teacherId);
        const timeKey = `${slot.start}-${slot.end}`;
        const idx = dayTimeIndexMap.get(slot.day)?.get(timeKey);
        if (Number.isInteger(idx) && st.dayIndices.has(slot.day)) {
          const daySet = st.dayIndices.get(slot.day);
          daySet.delete(idx);
          if (!daySet.size) {
            st.dayIndices.delete(slot.day);
            st.daySpan.delete(slot.day);
          } else {
            const span = st.daySpan.get(slot.day);
            if (span) {
              span.size -= 1;
              if (idx === span.min || idx === span.max) {
                let min = Infinity;
                let max = -Infinity;
                for (const value of daySet) {
                  if (value < min) min = value;
                  if (value > max) max = value;
                }
                span.min = min;
                span.max = max;
              }
            }
          }
        }
        const next = (st.dayCounts.get(slot.day) || 1) - 1;
        if (next <= 0) st.dayCounts.delete(slot.day); else st.dayCounts.set(slot.day, next);
      }

      function wouldExceedConsecutiveFast(teacherId, slot) {
        if (!maxConsecutive) return false;
        const st = getTeacherState(teacherId);
        const timeKey = `${slot.start}-${slot.end}`;
        const idx = dayTimeIndexMap.get(slot.day)?.get(timeKey);
        if (!Number.isInteger(idx)) return false;
        const daySet = st.dayIndices.get(slot.day);
        if (!daySet || !daySet.size) return false;
        let streak = 1;
        let left = idx - 1;
        let right = idx + 1;
        while (daySet.has(left)) { streak++; left--; if (streak > maxConsecutive) return true; }
        while (daySet.has(right)) { streak++; right++; if (streak > maxConsecutive) return true; }
        return streak > maxConsecutive;
      }

      function preferencePenaltyFast(teacherId, slot) {
        if (!avoidGaps && !balanceDaily) return 0;
        const st = getTeacherState(teacherId);
        let penalty = 0;

        if (avoidGaps) {
          const timeKey = `${slot.start}-${slot.end}`;
          const idx = dayTimeIndexMap.get(slot.day)?.get(timeKey);
          if (Number.isInteger(idx)) {
            const span = st.daySpan.get(slot.day);
            if (!span) {
              penalty += 0;
            } else {
              const oldGaps = Math.max(0, span.max - span.min + 1 - span.size);
              const newMin = Math.min(span.min, idx);
              const newMax = Math.max(span.max, idx);
              const newSize = span.size + (st.dayIndices.get(slot.day).has(idx) ? 0 : 1);
              const newGaps = Math.max(0, newMax - newMin + 1 - newSize);
              penalty += newGaps * 100;
              penalty += Math.max(0, newGaps - oldGaps) * 100;
            }
          }
        }

        if (balanceDaily) {
          const values = Array.from(st.dayCounts.values());
          const dayValue = (st.dayCounts.get(slot.day) || 0) + 1;
          values.push(dayValue);
          if (values.length > 1) penalty += (Math.max(...values) - Math.min(...values)) * 10;
        }

        return penalty;
      }

      for (const task of tasks) {
        const slotId = fixedAssignments.get(task.lessonId);
        if (!slotId) continue;
        const slot = slotById.get(slotId);
        if (!slot || slot.roomId !== task.roomId || isTeacherBlocked(task.teacherId, slot)) {
          generateSchedule.lastError = "fixed-invalid";
          return null;
        }
        const tk = `${task.teacherId}@@${slotKeyById.get(slot.id)}`;
        const rk = `${task.roomId}@@${slotId}`;
        if (usedTeacherSlot.has(tk) || usedRoomSlot.has(rk)) {
          generateSchedule.lastError = "fixed-conflict";
          return null;
        }
        assignment.set(task.lessonId, slotId);
        usedTeacherSlot.add(tk);
        usedRoomSlot.add(rk);
        addTeacherSlot(task.teacherId, slot);
      }

      const orderedTasks = tasks
        .filter((t) => !assignment.has(t.lessonId))
        .sort((a, b) => (a.teacherId === b.teacherId ? a.roomId.localeCompare(b.roomId) : a.teacherId.localeCompare(b.teacherId)));

      for (const task of orderedTasks) {
        const roomCandidates = roomSlotIds.get(task.roomId) || [];
        const blocked = blockedSlotsByTeacher.get(task.teacherId) || new Set();
        taskCandidates.set(task.lessonId, roomCandidates.filter((slotId) => !blocked.has(slotId)));
      }

      function slotOptions(task) {
        const out = [];
        const staticCandidates = taskCandidates.get(task.lessonId) || [];
        for (const slotId of staticCandidates) {
          const slot = slotById.get(slotId);
          const tk = `${task.teacherId}@@${slotKeyById.get(slotId)}`;
          const rk = `${task.roomId}@@${slotId}`;
          if (usedTeacherSlot.has(tk) || usedRoomSlot.has(rk)) continue;
          if (wouldExceedConsecutiveFast(task.teacherId, slot)) continue;
          out.push(slotId);
        }
        if (avoidGaps || balanceDaily) {
          out.sort((a, b) =>
            preferencePenaltyFast(task.teacherId, slotById.get(a)) -
            preferencePenaltyFast(task.teacherId, slotById.get(b))
          );
        }
        return out;
      }

      function backtrack(remaining) {
        steps++;
        if (steps > maxSteps || Date.now() > deadline) {
          aborted = true;
          return false;
        }
        if (!remaining.length) return true;
        let bestIdx = -1;
        let bestOps = null;
        for (let i = 0; i < remaining.length; i++) {
          const ops = slotOptions(remaining[i]);
          if (!ops.length) return false;
          if (!bestOps || ops.length < bestOps.length) {
            bestOps = ops;
            bestIdx = i;
            if (ops.length === 1) break;
          }
        }

        const [task] = remaining.splice(bestIdx, 1);
        for (const slotId of bestOps) {
          const slot = slotById.get(slotId);
          const tk = `${task.teacherId}@@${slotKeyById.get(slotId)}`;
          const rk = `${task.roomId}@@${slotId}`;
          usedTeacherSlot.add(tk);
          usedRoomSlot.add(rk);
          assignment.set(task.lessonId, slotId);
          addTeacherSlot(task.teacherId, slot);
          if (backtrack(remaining)) return true;
          assignment.delete(task.lessonId);
          usedTeacherSlot.delete(tk);
          usedRoomSlot.delete(rk);
          removeTeacherSlot(task.teacherId, slot);
          if (aborted) break;
        }
        remaining.splice(bestIdx, 0, task);
        return false;
      }

      const baseAssignment = new Map(assignment);
      const baseUsedTeacher = new Set(usedTeacherSlot);
      const baseUsedRoom = new Set(usedRoomSlot);
      const baseTeacherState = cloneTeacherStateMap(teacherState);

      if (!backtrack([...orderedTasks])) {
        if (!aborted) {
          generateSchedule.lastError = "no-solution";
          return null;
        }

        assignment.clear();
        usedTeacherSlot.clear();
        usedRoomSlot.clear();
        for (const [k, v] of baseAssignment.entries()) assignment.set(k, v);
        for (const v of baseUsedTeacher) usedTeacherSlot.add(v);
        for (const v of baseUsedRoom) usedRoomSlot.add(v);
        teacherState.clear();
        for (const [k, v] of baseTeacherState.entries()) teacherState.set(k, v);

        const remaining = [...orderedTasks];
        while (remaining.length) {
          let bestIdx = -1;
          let bestOps = null;
          for (let i = 0; i < remaining.length; i++) {
            const ops = slotOptions(remaining[i]);
            if (!ops.length) {
              generateSchedule.lastError = "timeout-no-solution";
              return null;
            }
            if (!bestOps || ops.length < bestOps.length) {
              bestOps = ops;
              bestIdx = i;
            }
          }

          const [task] = remaining.splice(bestIdx, 1);
          const slotId = bestOps[0];
          const slot = slotById.get(slotId);
          const tk = `${task.teacherId}@@${slotKeyById.get(slotId)}`;
          const rk = `${task.roomId}@@${slotId}`;
          assignment.set(task.lessonId, slotId);
          usedTeacherSlot.add(tk);
          usedRoomSlot.add(rk);
          addTeacherSlot(task.teacherId, slot);
        }
      }

      generateSchedule.lastError = "";
      return assignment;
    }

    function generateWithPreferences(tasks, fixedAssignments = new Map()) {
      const complexity = tasks.length + state.slots.length + state.teachers.length;
      const quickTimeLimit = Math.max(450, Math.min(1200, 350 + complexity * 2));
      const quickStepLimit = Math.max(30000, Math.min(120000, 18000 + complexity * 190));
      const deepTimeLimit = Math.max(1000, Math.min(3000, 700 + complexity * 5));
      const deepStepLimit = Math.max(90000, Math.min(320000, 60000 + complexity * 600));
      const base = {
        fixedAssignments,
        maxConsecutive: 2,
        avoidGaps: true,
        balanceDaily: true,
        timeLimitMs: quickTimeLimit,
        maxSteps: quickStepLimit
      };

      const attempts = [
        base,
        { ...base, timeLimitMs: deepTimeLimit, maxSteps: deepStepLimit },
        { ...base, maxConsecutive: null }
      ];

      for (const opts of attempts) {
        if (opts.maxConsecutive !== null && (!Number.isFinite(opts.maxConsecutive) || opts.maxConsecutive <= 0)) continue;
        const assignment = generateSchedule(tasks, opts);
        if (assignment) return { assignment, used: opts, relaxed: opts !== base };
      }
      return { assignment: null, used: base, relaxed: false };
    }

    function findLessonByTeacherAndSlot(teacherId, slotId) {
      return state.lessonInstances.find((l) => l.teacherId === teacherId && l.slotId === slotId) || null;
    }

    function findLessonByTeacherAndTime(teacherId, day, start, end) {
      const slotById = new Map(state.slots.map((slot) => [slot.id, slot]));
      return state.lessonInstances.find((lesson) => {
        if (lesson.teacherId !== teacherId || !lesson.slotId) return false;
        const slot = slotById.get(lesson.slotId);
        if (!slot) return false;
        return slot.day === day && slot.start === start && slot.end === end;
      }) || null;
    }

    function findLessonsByTeacherAndTime(teacherId, day, start, end) {
      const slotById = new Map(state.slots.map((slot) => [slot.id, slot]));
      return state.lessonInstances.filter((lesson) => {
        if (lesson.teacherId !== teacherId || !lesson.slotId) return false;
        const slot = slotById.get(lesson.slotId);
        if (!slot) return false;
        return slot.day === day && slot.start === start && slot.end === end;
      });
    }

    function findSlotByRoomAndTime(roomId, day, start, end) {
      return state.slots.find((slot) => slot.roomId === roomId && slot.day === day && slot.start === start && slot.end === end) || null;
    }

    function findRoomConflict(roomId, slotId, ignoreIds = []) {
      if (!slotId) return null;
      const ignored = new Set(ignoreIds);
      return state.lessonInstances.find((l) => l.roomId === roomId && l.slotId === slotId && !ignored.has(l.lessonId)) || null;
    }

    function findTeacherConflictAtTime(teacherId, slot, ignoreIds = []) {
      const ignored = new Set(ignoreIds);
      const timeKey = slotTimeKey(slot);
      const slotById = new Map(state.slots.map((item) => [item.id, item]));
      return state.lessonInstances.find((lesson) => {
        if (lesson.teacherId !== teacherId) return false;
        if (!lesson.slotId || ignored.has(lesson.lessonId)) return false;
        const lessonSlot = slotById.get(lesson.slotId);
        if (!lessonSlot) return false;
        return slotTimeKey(lessonSlot) === timeKey;
      }) || null;
    }

    function moveOrSwapLesson(draggedLessonId, targetTeacherId, targetSlotId) {
      const dragged = state.lessonInstances.find((l) => l.lessonId === draggedLessonId);
      if (!dragged) return;
      if (isLessonLocked(dragged.lessonId)) {
        setStatus("Aula travada. Clique para destravar antes de mover.", "error");
        return;
      }

      if (dragged.teacherId !== targetTeacherId) {
        setStatus("Arraste apenas na linha do mesmo professor.", "error");
        return;
      }

      const targetSlot = state.slots.find((s) => s.id === targetSlotId);
      if (!targetSlot || targetSlot.roomId !== dragged.roomId) {
        setStatus("Cada aula só pode ser alocada nos horários da sua própria série/sala.", "error");
        return;
      }

      const sourceSlotId = dragged.slotId || null;
      const targetLesson = findLessonByTeacherAndTime(targetTeacherId, targetSlot.day, targetSlot.start, targetSlot.end);
      if (targetLesson && targetLesson.lessonId === dragged.lessonId) return;
      if (targetLesson && isLessonLocked(targetLesson.lessonId)) {
        setStatus("Aula de destino está travada.", "error");
        return;
      }

      let targetBackSlotId = null;
      if (targetLesson && sourceSlotId) {
        const sourceSlot = state.slots.find((s) => s.id === sourceSlotId);
        if (!sourceSlot) {
          setStatus("Horário de origem inválido para troca.", "error");
          return;
        }
        const targetBackSlot = findSlotByRoomAndTime(targetLesson.roomId, sourceSlot.day, sourceSlot.start, sourceSlot.end);
        targetBackSlotId = targetBackSlot?.id || null;
      }

      if (targetLesson) targetLesson.slotId = targetBackSlotId;
      dragged.slotId = targetSlotId;

      rebuildScheduleFromLessons();
      clearStatus();
      renderAll();
    }

    function unscheduleLesson(lessonId) {
      const lesson = state.lessonInstances.find((l) => l.lessonId === lessonId);
      if (!lesson) return;
      if (isLessonLocked(lessonId)) {
        setStatus("Aula travada. Clique para destravar antes de remover.", "error");
        return;
      }
      lesson.slotId = null;
      delete state.lockedLessons[lessonId];
      rebuildScheduleFromLessons();
      clearStatus();
      renderAll();
    }

    function renderScheduleGlobal() {
      const wrapper = document.createElement("div");
      wrapper.className = "schedule-layout";
      const slotById = new Map(state.slots.map((slot) => [slot.id, slot]));
      const lessonsByTeacherAndTime = new Map();

      for (const lesson of state.lessonInstances) {
        if (!lesson.slotId) continue;
        const slot = slotById.get(lesson.slotId);
        if (!slot) continue;
        const key = `${lesson.teacherId}@@${slot.day}@@${slot.start}@@${slot.end}`;
        if (!lessonsByTeacherAndTime.has(key)) lessonsByTeacherAndTime.set(key, []);
        lessonsByTeacherAndTime.get(key).push(lesson);
      }

      const tableWrap = document.createElement("div");
      tableWrap.className = "table-wrap";

      const table = document.createElement("table");
      table.className = "schedule-compact";
      const thead = document.createElement("thead");
      const dayGroups = [];
      for (const day of Object.keys(dayOrder).sort((a, b) => dayOrder[a] - dayOrder[b])) {
        const uniqByTime = new Map();
        for (const slot of state.slots) {
          if (slot.day !== day) continue;
          const key = `${slot.start}-${slot.end}`;
          if (!uniqByTime.has(key)) uniqByTime.set(key, { day, start: slot.start, end: slot.end });
        }
        const lessons = Array.from(uniqByTime.values()).sort((a, b) => a.start.localeCompare(b.start) || a.end.localeCompare(b.end));
        if (lessons.length) dayGroups.push({ day, lessons });
      }

      const topRow = document.createElement("tr");
      const topLeft = document.createElement("th");
      topLeft.textContent = "Professor";
      topLeft.rowSpan = 2;
      topRow.appendChild(topLeft);

      for (const group of dayGroups) {
        const th = document.createElement("th");
        th.textContent = group.day;
        th.colSpan = group.lessons.length;
        th.classList.add("day-divider");
        topRow.appendChild(th);
      }
      thead.appendChild(topRow);

      const subRow = document.createElement("tr");
      for (const group of dayGroups) {
        for (let i = 0; i < group.lessons.length; i++) {
          const th = document.createElement("th");
          th.textContent = String(i + 1);
          th.title = `${group.day} ${group.lessons[i].start}-${group.lessons[i].end}`;
          if (i === group.lessons.length - 1) th.classList.add("day-divider");
          subRow.appendChild(th);
        }
      }
      thead.appendChild(subRow);
      table.appendChild(thead);

      const tbody = document.createElement("tbody");
      for (const teacher of state.teachers) {
        const tr = document.createElement("tr");
        const nameTd = document.createElement("td");
        nameTd.textContent = teacher.name;
        tr.appendChild(nameTd);

        for (const group of dayGroups) {
          for (let i = 0; i < group.lessons.length; i++) {
            const lessonTime = group.lessons[i];
            const td = document.createElement("td");
            td.className = "drop-cell";
            td.title = `${group.day} ${lessonTime.start}-${lessonTime.end}`;
            if (i === group.lessons.length - 1) td.classList.add("day-divider");

            td.addEventListener("dragover", (e) => {
              e.preventDefault();
              td.classList.add("drop-hover");
            });
            td.addEventListener("dragleave", () => td.classList.remove("drop-hover"));
            td.addEventListener("drop", (e) => {
              e.preventDefault();
              td.classList.remove("drop-hover");
              const lessonId = e.dataTransfer.getData("text/plain");
              if (!lessonId) return;
              const draggedLesson = state.lessonInstances.find((l) => l.lessonId === lessonId);
              if (!draggedLesson) return;
              const targetSlot = findSlotByRoomAndTime(draggedLesson.roomId, lessonTime.day, lessonTime.start, lessonTime.end);
              if (!targetSlot) {
                setStatus("Essa sala não possui esse horário cadastrado.", "error");
                return;
              }
              moveOrSwapLesson(lessonId, teacher.id, targetSlot.id);
            });

            const cellKey = `${teacher.id}@@${lessonTime.day}@@${lessonTime.start}@@${lessonTime.end}`;
            const lessons = lessonsByTeacherAndTime.get(cellKey) || [];
            if (lessons.length) {
              for (const lesson of lessons) {
                const chip = document.createElement("div");
                chip.className = "lesson-chip";
                const locked = isLessonLocked(lesson.lessonId);
                chip.classList.toggle("locked", locked);
                chip.classList.toggle("open", !locked);
                chip.draggable = !locked;
                chip.title = `Sala ${lesson.roomName}`;
                chip.addEventListener("click", () => {
                  if (locked) {
                    delete state.lockedLessons[lesson.lessonId];
                  } else {
                    state.lockedLessons[lesson.lessonId] = true;
                  }
                  saveState();
                  renderAll();
                });

                const text = document.createElement("span");
                text.textContent = lesson.roomName;
                chip.appendChild(text);

                chip.addEventListener("dragstart", (e) => {
                  chip.classList.add("dragging");
                  e.dataTransfer.setData("text/plain", lesson.lessonId);
                });
                chip.addEventListener("dragend", () => chip.classList.remove("dragging"));

                const remove = document.createElement("button");
                remove.type = "button";
                remove.textContent = "x";
                remove.title = "Excluir do horário";
                remove.disabled = locked;
                remove.addEventListener("click", (e) => {
                  e.stopPropagation();
                  unscheduleLesson(lesson.lessonId);
                });
                chip.appendChild(remove);
                td.appendChild(chip);
              }
            } else {
              const empty = document.createElement("span");
              empty.className = "empty-chip";
              empty.textContent = "Livre";
              td.appendChild(empty);
            }

            tr.appendChild(td);
          }
        }
        tbody.appendChild(tr);
      }
      table.appendChild(tbody);
      tableWrap.appendChild(table);
      wrapper.appendChild(tableWrap);

      const backlog = document.createElement("aside");
      backlog.className = "backlog";
      const h = document.createElement("h3");
      h.textContent = "Aulas fora do horário";
      backlog.appendChild(h);

      const byTeacher = new Map();
      for (const teacher of state.teachers) byTeacher.set(teacher.id, []);
      for (const lesson of state.lessonInstances) {
        if (!lesson.slotId) byTeacher.get(lesson.teacherId)?.push(lesson);
      }

      let hasUnscheduled = false;
      for (const teacher of state.teachers) {
        const lessons = byTeacher.get(teacher.id) || [];
        if (!lessons.length) continue;
        hasUnscheduled = true;

        const group = document.createElement("details");
        group.className = "backlog-group";
        const summary = document.createElement("summary");
        summary.textContent = `${teacher.name} (${lessons.length})`;
        group.appendChild(summary);

        const list = document.createElement("div");
        list.className = "backlog-items";
        for (const lesson of lessons) {
          const chip = document.createElement("div");
          chip.className = "lesson-chip";
          chip.draggable = true;
          chip.textContent = lesson.roomName;
          chip.title = `Professor ${teacher.name} - Sala ${lesson.roomName}`;
          chip.addEventListener("dragstart", (e) => {
            chip.classList.add("dragging");
            e.dataTransfer.setData("text/plain", lesson.lessonId);
          });
          chip.addEventListener("dragend", () => chip.classList.remove("dragging"));
          list.appendChild(chip);
        }
        group.appendChild(list);
        backlog.appendChild(group);
      }

      if (!hasUnscheduled) {
        const p = document.createElement("p");
        p.className = "muted";
        p.textContent = "Nenhuma aula pendente.";
        backlog.appendChild(p);
      }

      wrapper.appendChild(backlog);
      return wrapper;
    }

    function renderScheduleByRoom(roomId) {
      const room = state.rooms.find((r) => r.id === roomId);
      if (!room) {
        const p = document.createElement("p");
        p.className = "muted";
        p.textContent = "Selecione uma sala válida.";
        return p;
      }

      const table = document.createElement("table");
      const thead = document.createElement("thead");
      const hr = document.createElement("tr");
      const th1 = document.createElement("th"); th1.textContent = "Aula";
      const th2 = document.createElement("th"); th2.textContent = "Horário";
      const th3 = document.createElement("th"); th3.textContent = "Professor";
      hr.appendChild(th1); hr.appendChild(th2); hr.appendChild(th3);
      thead.appendChild(hr);
      table.appendChild(thead);

      const tbody = document.createElement("tbody");
      const grouped = groupedSlotsByDay()
        .map((g) => ({ day: g.day, slots: g.slots.filter((s) => s.roomId === room.id) }))
        .filter((g) => g.slots.length);

      for (const group of grouped) {
        const divider = document.createElement("tr");
        const dayTd = document.createElement("td");
        dayTd.textContent = group.day;
        dayTd.className = "group-title";
        dayTd.colSpan = 3;
        divider.appendChild(dayTd);
        tbody.appendChild(divider);

        for (let i = 0; i < group.slots.length; i++) {
          const slot = group.slots[i];
          const tr = document.createElement("tr");
          const tdLesson = document.createElement("td"); tdLesson.textContent = String(i + 1);
          const tdTime = document.createElement("td"); tdTime.textContent = `${slot.start}-${slot.end}`;
          const tdTeacher = document.createElement("td"); tdTeacher.textContent = state.schedule[room.id]?.[slot.id]?.teacherName || "Livre";
          tr.appendChild(tdLesson);
          tr.appendChild(tdTime);
          tr.appendChild(tdTeacher);
          tbody.appendChild(tr);
        }
      }
      table.appendChild(tbody);
      return table;
    }

    function renderSchedule() {
      scheduleOutput.innerHTML = "";
      if (!state.lessonInstances.length) {
        scheduleOutput.innerHTML = "<p class='muted' style='padding:10px;'>Gere o horário para visualizar a grade final.</p>";
        return;
      }
      if (viewMode.value === "room") {
        scheduleOutput.appendChild(renderScheduleByRoom(viewRoom.value || state.rooms[0]?.id));
      } else {
        scheduleOutput.appendChild(renderScheduleGlobal());
      }
    }

    function exportSchedulePdf() {
      if (!state.schedule) {
        setStatus("Gere o cronograma antes de exportar.", "error");
        return;
      }

      const table = scheduleOutput.querySelector("table");
      if (!table) {
        setStatus("Não há grade para exportar.", "error");
        return;
      }

      const popup = window.open("", "_blank");
      if (!popup) {
        setStatus("Permita pop-up para exportar o PDF.", "error");
        return;
      }

      const selectedRoomId = viewRoom.value || state.rooms[0]?.id || "";
      const title = viewMode.value === "room"
        ? `Horário por Sala - ${state.rooms.find((r) => r.id === selectedRoomId)?.name || ""}`.trim()
        : "Horário da Escola Inteira";

      popup.document.write(`<!doctype html>
<html lang=\"pt-BR\"><head><meta charset=\"UTF-8\" /><title>${title}</title>
<style>
body{font-family:Arial,sans-serif;padding:16px;color:#111827}
h1{font-size:18px;margin:0 0 12px}
table{width:100%;border-collapse:collapse;font-size:12px}
th,td{border:1px solid #d1d5db;padding:6px;text-align:left;vertical-align:top}
th{background:#f3f4f6}
.lesson-chip button{display:none!important}
</style>
</head><body><h1>${title}</h1>${table.outerHTML}</body></html>`);
      popup.document.close();
      popup.focus();
      popup.print();
    }

    function escapeCsvCell(value) {
      const raw = String(value ?? "");
      if (!/[;"\n\r,]/.test(raw)) return raw;
      return `"${raw.replace(/"/g, "\"\"")}"`;
    }

    function downloadCsv(filename, rows) {
      const csvContent = rows.map((row) => row.map(escapeCsvCell).join(";")).join("\r\n");
      const blob = new Blob([`\uFEFF${csvContent}`], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    }

    function currentGlobalDayGroups() {
      const dayGroups = [];
      for (const day of Object.keys(dayOrder).sort((a, b) => dayOrder[a] - dayOrder[b])) {
        const uniqByTime = new Map();
        for (const slot of state.slots) {
          if (slot.day !== day) continue;
          const key = `${slot.start}-${slot.end}`;
          if (!uniqByTime.has(key)) uniqByTime.set(key, { day, start: slot.start, end: slot.end });
        }
        const lessons = Array.from(uniqByTime.values()).sort((a, b) => a.start.localeCompare(b.start) || a.end.localeCompare(b.end));
        if (lessons.length) dayGroups.push({ day, lessons });
      }
      return dayGroups;
    }

    function exportScheduleCsv() {
      if (!state.schedule || !state.lessonInstances.length) {
        setStatus("Gere o cronograma antes de exportar CSV.", "error");
        return;
      }

      const stamp = new Date().toISOString().slice(0, 10);
      const slotById = new Map(state.slots.map((slot) => [slot.id, slot]));
      const rows = [];

      if (viewMode.value === "room") {
        const roomId = viewRoom.value || state.rooms[0]?.id;
        const room = state.rooms.find((r) => r.id === roomId);
        if (!room) {
          setStatus("Selecione uma sala válida para exportar CSV.", "error");
          return;
        }

        rows.push(["Horário por Sala", room.name, "", ""]);
        rows.push(["Dia", "Aula", "Horário", "Professor"]);
        const grouped = groupedSlotsByDay()
          .map((g) => ({ day: g.day, slots: g.slots.filter((s) => s.roomId === room.id) }))
          .filter((g) => g.slots.length);

        for (const group of grouped) {
          for (let i = 0; i < group.slots.length; i++) {
            const slot = group.slots[i];
            const teacherName = state.schedule[room.id]?.[slot.id]?.teacherName || "Livre";
            rows.push([group.day, String(i + 1), `${slot.start}-${slot.end}`, teacherName]);
          }
        }

        downloadCsv(`horario-sala-${room.name}-${stamp}.csv`, rows);
        setStatus("CSV da sala exportado com sucesso.", "ok");
        return;
      }

      const dayGroups = currentGlobalDayGroups();
      const dayHeader = ["Professor"];
      const lessonHeader = [""];
      for (const group of dayGroups) {
        for (let i = 0; i < group.lessons.length; i++) {
          const lesson = group.lessons[i];
          dayHeader.push(i === 0 ? group.day : "");
          lessonHeader.push(`Aula ${i + 1} (${lesson.start}-${lesson.end})`);
        }
      }
      rows.push(dayHeader);
      rows.push(lessonHeader);

      const lessonsByTeacherAndTime = new Map();
      for (const lesson of state.lessonInstances) {
        if (!lesson.slotId) continue;
        const slot = slotById.get(lesson.slotId);
        if (!slot) continue;
        const key = `${lesson.teacherId}@@${slot.day}@@${slot.start}@@${slot.end}`;
        if (!lessonsByTeacherAndTime.has(key)) lessonsByTeacherAndTime.set(key, []);
        lessonsByTeacherAndTime.get(key).push(lesson.roomName);
      }

      for (const teacher of state.teachers) {
        const row = [teacher.name];
        for (const group of dayGroups) {
          for (const lessonTime of group.lessons) {
            const key = `${teacher.id}@@${lessonTime.day}@@${lessonTime.start}@@${lessonTime.end}`;
            const rooms = lessonsByTeacherAndTime.get(key) || [];
            row.push(rooms.length ? rooms.join(" | ") : "Livre");
          }
        }
        rows.push(row);
      }

      downloadCsv(`horario-escola-${stamp}.csv`, rows);
      setStatus("CSV da grade escolar exportado com sucesso.", "ok");
    }

    function renderAll() {
      normalizeRoomData();
      normalizeLessonInstances();
      normalizeLockedLessons();
      rebuildScheduleFromLessons();
      renderRooms();
      renderSlotSelector();
      renderSelectors();
      renderLoads();
      renderBlocks();
      renderSchedule();
      renderAlerts();
      regenerateBtn.disabled = !state.lessonInstances.length;
      saveState();
    }

    roomForm.addEventListener("submit", (e) => {
      e.preventDefault();
      const input = document.getElementById("room-name");
      const name = input.value.trim().toUpperCase();
      if (!/^\d[A-Z]$/.test(name)) {
        setStatus("Sala inválida. Use número + letra (ex.: 1A).", "error");
        return;
      }
      if (state.rooms.some((r) => r.name === name)) {
        setStatus("Essa sala já existe.", "error");
        return;
      }
      const roomId = makeId("room");
      state.rooms.push({ id: roomId, name });
      sortRooms();
      input.value = "";
      state.lessonInstances = [];
      state.schedule = null;
      clearStatus();
      renderAll();
    });

    slotForm.addEventListener("submit", (e) => {
      e.preventDefault();
      const roomId = slotRoom.value;
      const day = document.getElementById("slot-day").value;
      const start = document.getElementById("slot-start").value;
      const end = document.getElementById("slot-end").value;
      const res = addSlot(roomId, day, start, end);
      if (!res.ok) {
        setStatus(res.reason === "duplicate" ? "Esse horário já existe." : res.reason, "error");
        return;
      }
      sortSlots();
      state.lessonInstances = [];
      state.schedule = null;
      clearStatus();
      slotForm.reset();
      renderAll();
    });

    slotBatchForm.addEventListener("submit", (e) => {
      e.preventDefault();
      const roomIds = state.rooms.map((r) => r.id);
      const selectedDays = Array.from(document.querySelectorAll("input[name='slot-batch-day']:checked")).map((el) => el.value);
      const start = document.getElementById("slot-batch-start").value;
      const duration = Math.floor(Number(document.getElementById("slot-batch-duration").value));
      const gap = Math.floor(Number(document.getElementById("slot-batch-break").value));
      const recessStartRaw = document.getElementById("slot-batch-recess-start").value;
      const recessEndRaw = document.getElementById("slot-batch-recess-end").value;
      const count = Math.floor(Number(document.getElementById("slot-batch-count").value));

      if (!roomIds.length) { setStatus("Cadastre ao menos uma sala antes de criar horários em lote.", "error"); return; }
      if (!selectedDays.length) { setStatus("Selecione ao menos um dia.", "error"); return; }
      if (!start) { setStatus("Informe o início da 1a aula.", "error"); return; }
      if (!Number.isFinite(duration) || duration <= 0 || !Number.isFinite(gap) || gap < 0 || !Number.isFinite(count) || count <= 0) {
        setStatus("Duração, troca e quantidade precisam ser válidos.", "error");
        return;
      }

      const hasRecess = Boolean(recessStartRaw || recessEndRaw);
      if ((recessStartRaw && !recessEndRaw) || (!recessStartRaw && recessEndRaw)) {
        setStatus("Preencha início e fim do intervalo, ou deixe ambos vazios.", "error");
        return;
      }

      let recessStart = -1;
      let recessEnd = -1;
      if (hasRecess) {
        recessStart = timeToMinutes(recessStartRaw);
        recessEnd = timeToMinutes(recessEndRaw);
        if (recessEnd <= recessStart) { setStatus("Intervalo inválido.", "error"); return; }
      }

      function simulateLessons(countPerDay) {
        const simulated = [];
        let cursor = timeToMinutes(start);
        for (let i = 0; i < countPerDay; i++) {
          if (hasRecess && cursor >= recessStart && cursor < recessEnd) cursor = recessEnd;
          let endCursor = cursor + duration;
          if (hasRecess && cursor < recessStart && endCursor > recessStart) {
            cursor = recessEnd;
            endCursor = cursor + duration;
          }
          simulated.push({ start: cursor, end: endCursor });
          cursor = endCursor + gap;
        }
        return simulated;
      }

      let added = 0;
      let duplicates = 0;

      for (const roomId of roomIds) {
        const simulated = simulateLessons(count);
        if (simulated.length && simulated[simulated.length - 1].end > 24 * 60) {
          setStatus("O lote ultrapassa 24:00. Ajuste os parâmetros.", "error");
          return;
        }
        for (const day of selectedDays) {
          for (const lesson of simulated) {
            const res = addSlot(roomId, day, minutesToTime(lesson.start), minutesToTime(lesson.end));
            if (res.ok) {
              added++;
            } else if (res.reason === "duplicate") {
              duplicates++;
            }
          }
        }
      }

      if (!added) {
        setStatus("Nenhum novo horário adicionado (todos já existiam).", "error");
        return;
      }

      sortSlots();
      state.lessonInstances = [];
      state.schedule = null;
      renderAll();
      clearStatus();
      setStatus(`Lote concluído: ${added} horário(s) adicionados${duplicates ? `, ${duplicates} duplicado(s) ignorado(s)` : ""}.`, "ok");
    });

    teacherForm.addEventListener("submit", (e) => {
      e.preventDefault();
      const input = document.getElementById("teacher-name");
      const name = input.value.trim();
      if (!name) return;
      if (state.teachers.some((t) => t.name.toLowerCase() === name.toLowerCase())) {
        setStatus("Esse professor já existe.", "error");
        return;
      }
      state.teachers.push({ id: makeId("teacher"), name });
      sortTeachers();
      input.value = "";
      state.lessonInstances = [];
      state.schedule = null;
      clearStatus();
      renderAll();
    });

    loadForm.addEventListener("submit", (e) => {
      e.preventDefault();
      if (!state.teachers.length || !state.rooms.length) {
        setStatus("Cadastre professores e salas antes de definir cargas.", "error");
        return;
      }
      const teacherId = loadTeacher.value;
      const roomId = loadRoom.value;
      const qty = Math.floor(Number(loadQty.value));
      if (!teacherId || !roomId || !Number.isFinite(qty) || qty <= 0) {
        setStatus("Informe professor, sala e quantidade válida.", "error");
        return;
      }
      state.loads[keyLoad(teacherId, roomId)] = qty;
      state.lessonInstances = [];
      state.schedule = null;
      clearStatus();
      renderAll();
    });

    blockForm.addEventListener("submit", (e) => {
      e.preventDefault();
      if (!state.teachers.length) {
        setStatus("Cadastre professor antes de bloquear horários.", "error");
        return;
      }

      const teacherId = blockTeacher.value;
      const day = blockDay.value;
      const fullDay = blockFullDay.checked;
      const start = blockStart.value;
      const end = blockEnd.value;

      if (!teacherId || !day) {
        setStatus("Informe professor e dia do bloqueio.", "error");
        return;
      }
      if (!fullDay && (!start || !end)) {
        setStatus("Informe professor, dia e horários do bloqueio.", "error");
        return;
      }
      if (!fullDay && end <= start) {
        setStatus("Fim do bloqueio deve ser maior que o início.", "error");
        return;
      }

      const blocks = state.teacherBlocks[teacherId] || [];
      const overlap = blocks.some((b) => {
        if (b.day !== day) return false;
        if (fullDay || b.fullDay) return true;
        return timeToMinutes(start) < timeToMinutes(b.end) && timeToMinutes(b.start) < timeToMinutes(end);
      });
      if (overlap) {
        setStatus("Já existe bloqueio sobreposto para esse professor.", "error");
        return;
      }

      blocks.push({ id: makeId("block"), day, start: fullDay ? "00:00" : start, end: fullDay ? "23:59" : end, fullDay });
      state.teacherBlocks[teacherId] = blocks;
      state.lessonInstances = [];
      state.schedule = null;
      clearStatus();
      renderAll();
    });

    removeRoomBtn.addEventListener("click", () => {
      const roomId = roomList.value;
      if (!roomId) return;
      removeEntity(state.rooms, roomId);
      state.slots = state.slots.filter((s) => s.roomId !== roomId);
      renderAll();
    });

    removeTeacherBtn.addEventListener("click", () => {
      const teacherId = teacherList.value;
      if (!teacherId) return;
      removeEntity(state.teachers, teacherId);
      delete state.teacherBlocks[teacherId];
      renderAll();
    });

    removeSlotBtn.addEventListener("click", () => {
      const slotId = slotSelect.value;
      if (!slotId) return;
      removeEntity(state.slots, slotId);
      renderAll();
    });

    buildForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      const tasks = buildLessonsFromLoads();
      const issue = canAllocate(tasks);
      if (issue) {
        state.lessonInstances = [];
        state.schedule = null;
        renderSchedule();
        runtimeAlertMessage = issue;
        setStatus(issue, "error");
        renderAlerts();
        saveState();
        return;
      }

      setGenerationBusy(true);
      showGenerationProgress("Gerando horário automaticamente");
      setStatus("Gerando cronograma, aguarde...", "ok");
      try {
        const result = await runGeneration(tasks, new Map(), (progress) => {
          updateGenerationProgress(progress.percent, progress.stage);
        });
        const assignment = result.assignment ? new Map(result.assignment) : null;
        if (!assignment) {
          state.lessonInstances = [];
          state.schedule = null;
          renderSchedule();
          const msg = "Não foi possível montar horário com os dados atuais.";
          runtimeAlertMessage = msg;
          setStatus(msg, "error");
          renderAlerts();
          saveState();
          return;
        }

        for (const lesson of tasks) {
          lesson.slotId = assignment.get(lesson.lessonId) || null;
        }

        state.lessonInstances = tasks;
        state.lockedLessons = {};
        runtimeAlertMessage = "";
        rebuildScheduleFromLessons();
        renderSchedule();
        updateGenerationProgress(100, "Concluido");
        setStatus(
          result.relaxed
            ? "Horário gerado com sucesso. Algumas preferências foram relaxadas para viabilizar a solução."
            : "Horário gerado com sucesso.",
          "ok"
        );
        renderAlerts();
        saveState();
      } catch (err) {
        if (err && err.message === "generation-cancelled") return;
        const msg = err && err.message === "generation-timeout"
          ? "A geração excedeu o tempo limite. Reduza cargas/bloqueios ou tente novamente."
          : "Falha ao gerar cronograma. Tente novamente.";
        runtimeAlertMessage = msg;
        setStatus(msg, "error");
        renderAlerts();
      } finally {
        hideGenerationProgress();
        setGenerationBusy(false);
      }
    });

    regenerateBtn.addEventListener("click", async () => {
      if (!state.lessonInstances.length) {
        setStatus("Gere o cronograma antes de regenerar.", "error");
        return;
      }
      const tasks = state.lessonInstances.map((lesson) => ({
        lessonId: lesson.lessonId,
        teacherId: lesson.teacherId,
        teacherName: lesson.teacherName,
        roomId: lesson.roomId,
        roomName: lesson.roomName,
        slotId: null
      }));

      const fixedAssignments = new Map();
      for (const lesson of state.lessonInstances) {
        if (!lesson.slotId || !isLessonLocked(lesson.lessonId)) continue;
        fixedAssignments.set(lesson.lessonId, lesson.slotId);
      }

      setGenerationBusy(true);
      showGenerationProgress("Gerando conforme opções");
      setStatus("Gerando cronograma, aguarde...", "ok");
      try {
        const result = await runGeneration(tasks, fixedAssignments, (progress) => {
          updateGenerationProgress(progress.percent, progress.stage);
        });
        const assignment = result.assignment ? new Map(result.assignment) : null;
        if (!assignment) {
          const msg = "Não foi possível regenerar mantendo as aulas travadas.";
          runtimeAlertMessage = msg;
          setStatus(msg, "error");
          renderAlerts();
          return;
        }

        for (const lesson of tasks) {
          lesson.slotId = assignment.get(lesson.lessonId) || null;
        }

        state.lessonInstances = tasks;
        normalizeLockedLessons();
        runtimeAlertMessage = "";
        rebuildScheduleFromLessons();
        renderSchedule();
        updateGenerationProgress(100, "Concluido");
        setStatus(
          result.relaxed
            ? "Cronograma regenerado. Algumas preferências foram relaxadas para viabilizar a solução."
            : "Cronograma regenerado com as preferências padrão.",
          "ok"
        );
        renderAlerts();
        saveState();
      } catch (err) {
        if (err && err.message === "generation-cancelled") return;
        const msg = err && err.message === "generation-timeout"
          ? "A regeneração excedeu o tempo limite. Ajuste as travas/cargas e tente novamente."
          : "Falha ao regenerar cronograma. Tente novamente.";
        runtimeAlertMessage = msg;
        setStatus(msg, "error");
        renderAlerts();
      } finally {
        hideGenerationProgress();
        setGenerationBusy(false);
      }
    });

    clearAllBtn.addEventListener("click", () => {
      const ok = window.confirm("Deseja limpar todos os dados da plataforma?");
      if (!ok) return;

      state.rooms = [];
      state.slots = [];
      state.teachers = [];
      state.loads = {};
      state.teacherBlocks = {};
      state.lockedLessons = {};
      state.lessonInstances = [];
      state.schedule = null;
      state.nextId = 1;
      runtimeAlertMessage = "";
      clearStatus();
      localStorage.removeItem(STORAGE_KEY);
      renderAll();
    });

    viewMode.addEventListener("change", () => {
      renderSelectors();
      renderSchedule();
      saveState();
    });

    viewRoom.addEventListener("change", () => {
      renderSchedule();
      saveState();
    });

    blockFullDay.addEventListener("change", () => {
      const disabled = blockFullDay.checked;
      blockStart.disabled = disabled || !state.teachers.length;
      blockEnd.disabled = disabled || !state.teachers.length;
      if (disabled) {
        blockStart.value = "";
        blockEnd.value = "";
      }
    });

    exportPdfBtn.addEventListener("click", exportSchedulePdf);
    exportCsvBtn.addEventListener("click", exportScheduleCsv);

    loadState();
    renderAll();


