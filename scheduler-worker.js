"use strict";

function timeToMinutes(value) {
  const [h, m] = String(value || "00:00").split(":").map(Number);
  return h * 60 + m;
}

function slotTimeKey(slot) {
  return `${slot.day}@@${slot.start}@@${slot.end}`;
}

function isTeacherBlocked(teacherBlocks, teacherId, slot) {
  const blocks = teacherBlocks?.[teacherId] || [];
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

function createFastStructures(ctx, tasks, options = {}) {
  const { slots, teachers, teacherBlocks } = ctx;
  const slotById = new Map(slots.map((s) => [s.id, s]));
  const dayTimeIndexMap = buildDayTimeIndexMap(slotById);
  const fixedAssignments = options.fixedAssignments instanceof Map ? options.fixedAssignments : new Map();
  const maxConsecutive = Number.isInteger(options.maxConsecutive) && options.maxConsecutive > 0 ? options.maxConsecutive : null;
  const avoidGaps = Boolean(options.avoidGaps);
  const balanceDaily = Boolean(options.balanceDaily);
  const slotKeyById = new Map();
  const roomSlotIds = new Map();
  const blockedSlotsByTeacher = new Map();
  const teacherIds = new Set();

  for (const slot of slots) {
    slotKeyById.set(slot.id, slotTimeKey(slot));
    if (!roomSlotIds.has(slot.roomId)) roomSlotIds.set(slot.roomId, []);
    roomSlotIds.get(slot.roomId).push(slot.id);
  }
  for (const teacher of teachers) teacherIds.add(teacher.id);
  for (const task of tasks) teacherIds.add(task.teacherId);
  for (const teacherId of teacherIds) {
    const blocked = new Set();
    for (const slot of slots) {
      if (isTeacherBlocked(teacherBlocks, teacherId, slot)) blocked.add(slot.id);
    }
    blockedSlotsByTeacher.set(teacherId, blocked);
  }

  return {
    slotById,
    dayTimeIndexMap,
    fixedAssignments,
    maxConsecutive,
    avoidGaps,
    balanceDaily,
    slotKeyById,
    roomSlotIds,
    blockedSlotsByTeacher
  };
}

function solveGreedyFast(ctx, tasks, options = {}) {
  const structures = createFastStructures(ctx, tasks, options);
  const {
    slotById,
    dayTimeIndexMap,
    fixedAssignments,
    maxConsecutive,
    avoidGaps,
    balanceDaily,
    slotKeyById,
    roomSlotIds,
    blockedSlotsByTeacher
  } = structures;

  const hardDeadline = Number.isFinite(options.absoluteDeadline) ? options.absoluteDeadline : Infinity;
  const usedTeacherSlot = new Set();
  const usedRoomSlot = new Set();
  const assignment = new Map();
  const teacherState = new Map();

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
      if (!span) st.daySpan.set(slot.day, { min: idx, max: idx, size: 1 });
      else {
        span.min = Math.min(span.min, idx);
        span.max = Math.max(span.max, idx);
        span.size += 1;
      }
    }
    st.dayCounts.set(slot.day, (st.dayCounts.get(slot.day) || 0) + 1);
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
        if (span) {
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
    if (Date.now() > hardDeadline) return null;
    const slotId = fixedAssignments.get(task.lessonId);
    if (!slotId) continue;
    const slot = slotById.get(slotId);
    if (!slot || slot.roomId !== task.roomId || isTeacherBlocked(ctx.teacherBlocks, task.teacherId, slot)) return null;
    const tk = `${task.teacherId}@@${slotKeyById.get(slot.id)}`;
    const rk = `${task.roomId}@@${slotId}`;
    if (usedTeacherSlot.has(tk) || usedRoomSlot.has(rk)) return null;
    assignment.set(task.lessonId, slotId);
    usedTeacherSlot.add(tk);
    usedRoomSlot.add(rk);
    addTeacherSlot(task.teacherId, slot);
  }

  const remaining = tasks
    .filter((t) => !assignment.has(t.lessonId))
    .map((task) => {
      const roomCandidates = roomSlotIds.get(task.roomId) || [];
      const blocked = blockedSlotsByTeacher.get(task.teacherId) || new Set();
      return {
        task,
        candidates: roomCandidates.filter((slotId) => !blocked.has(slotId))
      };
    })
    .sort((a, b) => a.candidates.length - b.candidates.length);

  for (const item of remaining) {
    if (Date.now() > hardDeadline) return null;
    if (!item.candidates.length) return null;
    let bestSlotId = null;
    let bestScore = Infinity;
    for (const slotId of item.candidates) {
      if (Date.now() > hardDeadline) return null;
      const slot = slotById.get(slotId);
      const tk = `${item.task.teacherId}@@${slotKeyById.get(slotId)}`;
      const rk = `${item.task.roomId}@@${slotId}`;
      if (usedTeacherSlot.has(tk) || usedRoomSlot.has(rk)) continue;
      if (wouldExceedConsecutiveFast(item.task.teacherId, slot)) continue;
      const score = preferencePenaltyFast(item.task.teacherId, slot);
      if (score < bestScore) {
        bestScore = score;
        bestSlotId = slotId;
      }
    }
    if (!bestSlotId) return null;
    const slot = slotById.get(bestSlotId);
    const tk = `${item.task.teacherId}@@${slotKeyById.get(bestSlotId)}`;
    const rk = `${item.task.roomId}@@${bestSlotId}`;
    assignment.set(item.task.lessonId, bestSlotId);
    usedTeacherSlot.add(tk);
    usedRoomSlot.add(rk);
    addTeacherSlot(item.task.teacherId, slot);
  }

  return assignment;
}

function generateSchedule(ctx, tasks, options = {}) {
  const { slots, teachers, teacherBlocks } = ctx;
  const slotById = new Map(slots.map((s) => [s.id, s]));
  const dayTimeIndexMap = buildDayTimeIndexMap(slotById);
  const fixedAssignments = options.fixedAssignments instanceof Map ? options.fixedAssignments : new Map();
  const maxConsecutive = Number.isInteger(options.maxConsecutive) && options.maxConsecutive > 0 ? options.maxConsecutive : null;
  const avoidGaps = Boolean(options.avoidGaps);
  const balanceDaily = Boolean(options.balanceDaily);
  const deadline = Date.now() + (Number.isFinite(options.timeLimitMs) ? options.timeLimitMs : 1800);
  const hardDeadline = Math.min(
    deadline,
    Number.isFinite(options.absoluteDeadline) ? options.absoluteDeadline : Infinity
  );
  const maxSteps = Number.isFinite(options.maxSteps) ? options.maxSteps : 350000;
  const slotKeyById = new Map();
  const roomSlotIds = new Map();
  const blockedSlotsByTeacher = new Map();
  const teacherIds = new Set();

  for (const slot of slots) {
    slotKeyById.set(slot.id, slotTimeKey(slot));
    if (!roomSlotIds.has(slot.roomId)) roomSlotIds.set(slot.roomId, []);
    roomSlotIds.get(slot.roomId).push(slot.id);
  }
  for (const teacher of teachers) teacherIds.add(teacher.id);
  for (const task of tasks) teacherIds.add(task.teacherId);
  for (const teacherId of teacherIds) {
    const blocked = new Set();
    for (const slot of slots) {
      if (isTeacherBlocked(teacherBlocks, teacherId, slot)) blocked.add(slot.id);
    }
    blockedSlotsByTeacher.set(teacherId, blocked);
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
        if (span) {
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
    if (Date.now() > hardDeadline) return null;
    const slotId = fixedAssignments.get(task.lessonId);
    if (!slotId) continue;
    const slot = slotById.get(slotId);
    if (!slot || slot.roomId !== task.roomId || isTeacherBlocked(teacherBlocks, task.teacherId, slot)) return null;
    const tk = `${task.teacherId}@@${slotKeyById.get(slot.id)}`;
    const rk = `${task.roomId}@@${slotId}`;
    if (usedTeacherSlot.has(tk) || usedRoomSlot.has(rk)) return null;
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
    if (Date.now() > hardDeadline) return [];
    const out = [];
    const staticCandidates = taskCandidates.get(task.lessonId) || [];
      for (const slotId of staticCandidates) {
      if (Date.now() > hardDeadline) return [];
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
    if (steps > maxSteps || Date.now() > hardDeadline) {
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
    if (!aborted) return null;

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
      if (Date.now() > hardDeadline) return null;
      let bestIdx = -1;
      let bestOps = null;
      for (let i = 0; i < remaining.length; i++) {
        if (Date.now() > hardDeadline) return null;
        const ops = slotOptions(remaining[i]);
        if (!ops.length) return null;
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

  return assignment;
}

function generateWithPreferences(payload, onProgress) {
  const tasks = payload.tasks || [];
  const fixedAssignments = new Map(payload.fixedAssignments || []);
  const slots = payload.slots || [];
  const teachers = payload.teachers || [];
  const teacherBlocks = payload.teacherBlocks || {};
  const ctx = { slots, teachers, teacherBlocks };
  const complexity = tasks.length + slots.length + teachers.length;
  const quickTimeLimit = Math.max(450, Math.min(1200, 350 + complexity * 2));
  const quickStepLimit = Math.max(30000, Math.min(120000, 18000 + complexity * 190));
  const deepTimeLimit = Math.max(1000, Math.min(3000, 700 + complexity * 5));
  const deepStepLimit = Math.max(90000, Math.min(320000, 60000 + complexity * 600));
  const globalTimeLimit = Math.max(2500, Math.min(9000, 1800 + complexity * 8));
  const absoluteDeadline = Date.now() + globalTimeLimit;
  const base = {
    fixedAssignments,
    maxConsecutive: 2,
    avoidGaps: true,
    balanceDaily: true,
    timeLimitMs: quickTimeLimit,
    maxSteps: quickStepLimit,
    absoluteDeadline
  };

  const attempts = [
    base,
    { ...base, timeLimitMs: deepTimeLimit, maxSteps: deepStepLimit },
    { ...base, maxConsecutive: null }
  ];

  if (typeof onProgress === "function") {
    onProgress({ percent: 5, stage: "Iniciando busca de alocacao" });
  }

  const fastFirst = solveGreedyFast(ctx, tasks, base);
  if (fastFirst) {
    if (typeof onProgress === "function") onProgress({ percent: 72, stage: "Heuristica rapida concluida" });
    return {
      assignment: Array.from(fastFirst.entries()),
      used: { ...base, fixedAssignments: undefined },
      relaxed: false
    };
  }

  for (let i = 0; i < attempts.length; i++) {
    if (Date.now() > absoluteDeadline) return { assignment: null, used: { ...base, fixedAssignments: undefined }, relaxed: true };
    const opts = attempts[i];
    if (typeof onProgress === "function") {
      const ratio = (i + 1) / attempts.length;
      onProgress({
        percent: Math.round(8 + ratio * 84),
        stage: `Testando estrategia ${i + 1}/${attempts.length}`
      });
    }
    if (opts.maxConsecutive !== null && (!Number.isFinite(opts.maxConsecutive) || opts.maxConsecutive <= 0)) continue;
    const assignment = generateSchedule(ctx, tasks, opts);
    if (assignment) {
      if (typeof onProgress === "function") onProgress({ percent: 98, stage: "Consolidando resultado" });
      return {
        assignment: Array.from(assignment.entries()),
        used: { ...opts, fixedAssignments: undefined },
        relaxed: opts !== base
      };
    }
  }
  return { assignment: null, used: { ...base, fixedAssignments: undefined }, relaxed: false };
}

self.addEventListener("message", (event) => {
  const data = event.data || {};
  if (data.type !== "generate") return;
  try {
    const result = generateWithPreferences(data, (progress) => {
      self.postMessage({ type: "progress", requestId: data.requestId, progress });
    });
    self.postMessage({ type: "result", requestId: data.requestId, result });
  } catch (error) {
    self.postMessage({
      type: "error",
      requestId: data.requestId,
      error: error && error.message ? error.message : "Erro interno no worker."
    });
  }
});
