const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbynJbAbOoHOIcB4lYlb5gdHvGOe-hXnMNOMYrKy9fh9mv-5K7-nZYx9GiaBYkq8_yWFRg/exec";

const STANDARD_SLOT_TIMES = [
  "9:00 AM - 9:20 AM",
  "9:20 AM - 9:40 AM",
  "9:40 AM - 10:00 AM",
  "10:00 AM - 10:20 AM",
  "10:20 AM - 10:40 AM",
  "10:40 AM - 11:00 AM",
  "11:00 AM - 11:20 AM",
  "11:20 AM - 11:40 AM",
  "11:40 AM - 12:00 PM",
];

const availability = [
  { date: "2026-08-05", times: ["9:00 AM - 9:20 AM", "9:20 AM - 9:40 AM", "9:40 AM - 10:00 AM", "1:00 PM - 1:20 PM", "1:20 PM - 1:40 PM", "1:40 PM - 2:00 PM"], capacity: 1 },
  { date: "2026-08-06", times: STANDARD_SLOT_TIMES, capacity: 5 },
  { date: "2026-08-07", times: STANDARD_SLOT_TIMES, capacity: 4 },
];

const availabilityMap = new Map(availability.map((entry) => [entry.date, entry]));

// Slots forced to display as already booked, with no real booking behind them.
const FORCED_BOOKED_SLOTS = new Set([
  "Wednesday, August 5, 2026 at 1:20 PM - 1:40 PM",
  "Wednesday, August 5, 2026 at 1:40 PM - 2:00 PM",
]);

const form = document.getElementById("interview-form");
const submitBtn = document.getElementById("submit-btn");
const errorMessage = document.getElementById("error-message");
const confirmation = document.getElementById("confirmation");
const confirmedSlot = document.getElementById("confirmed-slot");
const calendarLoading = document.getElementById("calendar-loading");
const calendarEl = document.getElementById("calendar");
const calGrid = document.getElementById("calendar-grid");
const calMonthLabel = document.getElementById("cal-month-label");
const calPrevBtn = document.getElementById("cal-prev");
const calNextBtn = document.getElementById("cal-next");
const timeField = document.getElementById("time-field");
const timePicker = document.getElementById("time-picker");
const timeSlotInput = document.getElementById("timeSlot");

let bookedCounts = new Map();
let selectedDateISO = null;
let currentMonth = parseISODate(availability[0].date);
currentMonth.setDate(1);

function parseISODate(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d);
}

function toISODate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function formatDateLabel(date) {
  return date.toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" });
}

function hasAvailabilityInMonth(year, month, comparator) {
  return availability.some((entry) => {
    const d = parseISODate(entry.date);
    return comparator(d.getFullYear(), d.getMonth(), year, month);
  });
}

function renderCalendar() {
  const year = currentMonth.getFullYear();
  const month = currentMonth.getMonth();

  calMonthLabel.textContent = currentMonth.toLocaleDateString("en-US", { month: "long", year: "numeric" });

  const firstOfMonth = new Date(year, month, 1);
  const startWeekday = firstOfMonth.getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  calGrid.innerHTML = "";

  for (let i = 0; i < startWeekday; i++) {
    const blank = document.createElement("div");
    blank.className = "cal-cell cal-blank";
    calGrid.appendChild(blank);
  }

  for (let day = 1; day <= daysInMonth; day++) {
    const cellDate = new Date(year, month, day);
    const iso = toISODate(cellDate);
    const isAvailable = availabilityMap.has(iso);

    const cell = document.createElement("button");
    cell.type = "button";
    cell.className = "cal-cell";
    cell.textContent = String(day);

    if (isAvailable) {
      cell.classList.add("cal-available");
      if (selectedDateISO === iso) cell.classList.add("selected");
      cell.addEventListener("click", () => selectDate(iso, cellDate));
    } else {
      cell.disabled = true;
      cell.classList.add("cal-disabled");
    }

    calGrid.appendChild(cell);
  }

  calPrevBtn.disabled = !hasAvailabilityInMonth(
    year, month,
    (dy, dm, y, m) => dy < y || (dy === y && dm < m)
  );
  calNextBtn.disabled = !hasAvailabilityInMonth(
    year, month,
    (dy, dm, y, m) => dy > y || (dy === y && dm > m)
  );
}

function selectDate(iso, dateObj) {
  selectedDateISO = iso;
  timeSlotInput.value = "";
  renderCalendar();
  renderTimeSlots(iso, dateObj);
}

function renderTimeSlots(iso, dateObj) {
  const entry = availabilityMap.get(iso);
  const times = entry ? entry.times : [];
  const capacity = entry ? entry.capacity : 1;
  const dateLabel = formatDateLabel(dateObj);
  timePicker.innerHTML = "";

  times.forEach((time) => {
    const slotLabel = `${dateLabel} at ${time}`;
    const bookedCount = bookedCounts.get(slotLabel) || 0;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "picker-btn";

    if (bookedCount >= capacity || FORCED_BOOKED_SLOTS.has(slotLabel)) {
      btn.classList.add("booked");
      btn.disabled = true;
      btn.textContent = `${time} (Booked)`;
    } else {
      btn.textContent = time;
      btn.addEventListener("click", () => selectTime(btn, slotLabel));
    }

    timePicker.appendChild(btn);
  });

  timeField.hidden = false;
}

function selectTime(button, slotLabel) {
  [...timePicker.children].forEach((btn) => btn.classList.remove("selected"));
  button.classList.add("selected");
  timeSlotInput.value = slotLabel;
}

async function loadBookedSlots() {
  try {
    const res = await fetch(SCRIPT_URL);
    const data = await res.json();
    bookedCounts = new Map();
    (data.bookedSlots || []).forEach((slotLabel) => {
      bookedCounts.set(slotLabel, (bookedCounts.get(slotLabel) || 0) + 1);
    });
  } catch (err) {
    console.warn("Could not load booked slots; showing all as available.", err);
    bookedCounts = new Map();
  }
}

calPrevBtn.addEventListener("click", () => {
  currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() - 1, 1);
  renderCalendar();
});

calNextBtn.addEventListener("click", () => {
  currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1);
  renderCalendar();
});

(async function init() {
  await loadBookedSlots();
  calendarLoading.hidden = true;
  calendarEl.hidden = false;
  renderCalendar();
})();

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  errorMessage.hidden = true;

  const payload = {
    name: form.name.value.trim(),
    email: form.email.value.trim(),
    phone: form.phone.value.trim(),
    timeSlot: timeSlotInput.value,
  };

  if (!payload.name || !payload.email || !payload.phone || !payload.timeSlot) {
    errorMessage.textContent = "Please fill out every field and pick a date and time before submitting.";
    errorMessage.hidden = false;
    return;
  }

  submitBtn.disabled = true;
  submitBtn.textContent = "Submitting...";

  try {
    const res = await fetch(SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload),
    });
    const result = await res.json();

    if (result.status === "success") {
      confirmedSlot.textContent = payload.timeSlot;
      form.hidden = true;
      confirmation.hidden = false;
    } else {
      errorMessage.textContent = result.message || "That slot may already be booked. Please choose another time.";
      errorMessage.hidden = false;
      timeSlotInput.value = "";
      await loadBookedSlots();
      if (selectedDateISO) {
        renderTimeSlots(selectedDateISO, parseISODate(selectedDateISO));
      }
    }
  } catch (err) {
    errorMessage.textContent = "Something went wrong submitting your request. Please try again.";
    errorMessage.hidden = false;
  } finally {
    submitBtn.disabled = false;
    submitBtn.textContent = "Schedule Interview";
  }
});
