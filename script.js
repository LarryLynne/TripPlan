// Хранилища данных
window.allTrips = [];
window.allVehicles = [];
window.currentFilteredTrips = [];    // ТІЛЬКИ ТЕ, ЩО НА ЕКРАНІ
window.currentFilteredVehicles = []; // ТІЛЬКИ ТІ МАШИНИ, ЩО ЗЛІВА
let nodeDictionary = new Map();
let clusterizeInstance = null;
let currentTeleportCar = null;

// Инициализация интерфейса
const dropZoneTrips = document.getElementById('drop_zone');
const dropZoneCars = document.getElementById('drop_zone_cars');
const fileInput = document.getElementById('file_input');
const carsInput = document.getElementById('cars_input');

// Прив'язка подій до всіх фільтрів для миттєвої реакції
document.getElementById('node_filter').addEventListener('input', applyFilters);
document.getElementById('filter_car_number').addEventListener('input', applyFilters);
document.querySelectorAll('.trip-filter').forEach(el => el.addEventListener('input', applyFilters));
document.querySelectorAll('.trip-filter').forEach(el => el.addEventListener('change', applyFilters));

// Універсальна функція застосування фільтрів
function applyFilters() {
    const nodeFilter = (document.getElementById('node_filter').value || "").trim().toLowerCase();
    
    // Фільтри графіків
    const dayFilter = document.getElementById('filter_day').value;
    const statusFilter = document.getElementById('filter_status').value;
    const autoTypeFilter = (document.getElementById('filter_auto_type').value || "").trim().toLowerCase();
    const tripTypeFilter = (document.getElementById('filter_trip_type').value || "").trim().toLowerCase();
    const assignedCarFilter = (document.getElementById('filter_assigned_car').value || "").trim().toLowerCase();
    const destFilter = (document.getElementById('filter_destination').value || "").trim().toLowerCase();

    // 1. Фільтруємо графіки і ЗБЕРІГАЄМО в глобальну змінну
    window.currentFilteredTrips = window.allTrips.filter(t => {
        if (nodeFilter && !t.origin.toLowerCase().includes(nodeFilter)) return false;
        
        // НОВИЙ РЯДОК: Перевіряємо отримувача
        if (destFilter && !t.destination.toLowerCase().includes(destFilter)) return false;
        
        if (dayFilter !== "" && t.logDays[parseInt(dayFilter)] !== '+') return false;
        if (statusFilter === 'with_auto' && !t.assignedCar) return false;
        if (statusFilter === 'without_auto' && t.assignedCar) return false;
        if (autoTypeFilter && !(t.auto || "").toLowerCase().includes(autoTypeFilter)) return false;
        if (tripTypeFilter && !(t.type || "").toLowerCase().includes(tripTypeFilter)) return false;
        if (assignedCarFilter && !(t.assignedCar || "").toLowerCase().includes(assignedCarFilter)) return false;
        return true; 
    });

    // 2. Фільтруємо машини і теж ЗБЕРІГАЄМО
    const carNumberFilter = (document.getElementById('filter_car_number').value || "").trim().toLowerCase();
    window.currentFilteredVehicles = window.allVehicles.filter(v => {
        if (nodeFilter && !v.node.toLowerCase().includes(nodeFilter)) return false;
        if (carNumberFilter && !v.number.toLowerCase().includes(carNumberFilter)) return false;
        return true;
    });

    // Відмальовуємо вже відфільтровані масиви
    render(window.currentFilteredTrips);
    renderVehicles(); 
}

// Слушатели завантаження файлів
dropZoneTrips.onclick = () => fileInput.click();
dropZoneCars.onclick = () => carsInput.click();

fileInput.onchange = e => handleFile(e.target.files[0], 'trips');
carsInput.onchange = e => handleFile(e.target.files[0], 'cars');

function setupFileDragAndDrop(zone, type) {
    zone.addEventListener('dragover', e => {
        e.preventDefault(); // Обов'язково, щоб браузер дозволив "кинути" файл
        e.stopPropagation();
        // Візуальний відгук: підсвічуємо зону при наведенні
        zone.style.backgroundColor = type === 'cars' ? '#e8f5e9' : '#e8f0fe';
    });

    zone.addEventListener('dragleave', e => {
        e.preventDefault();
        e.stopPropagation();
        zone.style.backgroundColor = ''; // Прибираємо підсвічування
    });

    zone.addEventListener('drop', e => {
        e.preventDefault();
        e.stopPropagation();
        zone.style.backgroundColor = '';
        
        // Беремо перший перетягнутий файл і передаємо його нашій функції
        const files = e.dataTransfer.files;
        if (files && files.length > 0) {
            handleFile(files[0], type);
        }
    });
}

// Активуємо Drag-and-Drop для обох зон
setupFileDragAndDrop(dropZoneTrips, 'trips');
setupFileDragAndDrop(dropZoneCars, 'cars');

async function handleFile(file, type) {
    if (!file) return;
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });

    if (type === 'trips') {
        window.allTrips = rows.slice(4).filter(r => r[2]).map((r, idx) => {
            let t = new Trip(r);
            t.id = `trip_${idx}_${Date.now()}`;
            t.assignedCar = null;
            return t;
        });

        // НОВЕ: Сортування графіків
        window.allTrips.sort((a, b) => {
            // 1. Спочатку жорстко групуємо за логістичним днем
            if (a.logisticDay !== b.logisticDay) {
                return a.logisticDay - b.logisticDay;
            }
            // 2. А вже всередині одного логістичного дня шикуємо за часом виїзду
            return a.depInt - b.depInt;
        });

        applyFilters();
    } else {
        window.allVehicles = rows.slice(1).filter(r => r[0]).map(r => {
            const initTime = parseExcelTime(r[2]); // Зберігаємо час
            return {
                number: String(r[0]),
                type: String(r[1]),
                availableTime: initTime,
                node: String(r[3]).trim(),
                city: r[5],
                originalNode: String(r[3]).trim(),
                originalTime: initTime // <-- ДОДАЛИ ЦЕЙ РЯДОК
            };
        });
        applyFilters();
    }
}

function parseExcelTime(val) {
    if (typeof val === 'string') {
        const days = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Нд'];
        let d = 0;
        days.forEach((day, idx) => { if (val.includes(day)) d = idx; });
        const timeMatch = val.match(/(\d{1,2}):(\d{2})/);
        if (timeMatch) {
            return d * 1440 + parseInt(timeMatch[1]) * 60 + parseInt(timeMatch[2]);
        }
        return 0;
    }
    if (typeof val === 'number') {
        if (val < 1) {
            return Math.round(val * 1440);
        } else {
            let daysSince1900 = Math.floor(val);
            let dayOfWeek = (daysSince1900 - 2) % 7; 
            if (dayOfWeek < 0) dayOfWeek += 7; 
            let timeFraction = val - daysSince1900;
            let mins = Math.round(timeFraction * 1440);
            return (dayOfWeek * 1440) + mins;
        }
    }
    return 0;
}

class Trip {
    constructor(r) {
        this.grf = r[2]; 
        this.digit = r[3]; 
        this.code = r[4]; 
        this.group = r[5];
        this.naryad = r[6]; 
        this.type = r[7]; 
        
        let rawAuto = String(r[8] || "").trim();
        if (rawAuto.toUpperCase() === "БДФ") {
            this.auto = "Шасі BDF";
        } else {
            this.auto = r[8];
        }

        this.load = r[9];
        this.route = String(r[57] || "").trim();
        const routeParts = this.route.split('_');
        this.origin = routeParts.length > 0 ? routeParts[0].trim() : "";
        this.destination = routeParts.length > 1 ? routeParts[routeParts.length - 1].trim() : "";

        this.originData = nodeDictionary.get(this.origin) || { city: this.origin, city2: this.origin };
        this.destData = nodeDictionary.get(this.destination) || { city: this.destination, city2: this.destination };

        const astroDays = [r[16], r[17], r[18], r[19], r[20], r[21], r[22]];
        this.dayIndex = astroDays.findIndex(d => d === '+');
        if (this.dayIndex === -1) this.dayIndex = 0;

        this.logDays = [r[23], r[24], r[25], r[26], r[27], r[28], r[29]];
        this.logisticDay = this.logDays.findIndex(d => d === "+");
        if (this.logisticDay === -1) this.logisticDay = 0;

        this.drivers = r[30]; 
        this.deadline = r[31];
        
        this.podachaStr = formatTimeStr(r[32]);
        this.depStr = formatTimeStr(r[33]);
        this.arrStr = formatTimeStr(r[40]);
        this.freeStr = formatTimeStr(r[41]);

        this.calculateTimeline();
        this.calculateTrueTimes();
        
        this.comment = r[54];
        this.ringId = null;
        this.assignedCar = null; 
    }

    calculateTrueTimes() {
        const autoStr = String(this.auto || "").toUpperCase();
        const isBDF = autoStr.includes("БДФ") || autoStr.includes("BDF");
        this.trueStart = isBDF ? this.depInt : this.podachaInt;
        this.trueEnd = isBDF ? this.arrInt : this.freeInt;
    }

    calculateTimeline() {
        const minInDay = 1440; const minInWeek = 10080;
        const dayStart = this.dayIndex * minInDay;
        const toMin = (str) => {
            if (!str) return 0;
            const [h, m] = str.split(':').map(Number);
            return (h * 60) + m;
        };
        const dM = toMin(this.depStr);
        this.depInt = dayStart + dM;
        let pM = toMin(this.podachaStr);
        let pInt = dayStart + pM;
        if (pM > dM) pInt -= minInDay;
        this.podachaInt = pInt < 0 ? pInt + minInWeek : pInt;
        let aM = toMin(this.arrStr);
        let aInt = dayStart + aM;
        if (aM < dM) aInt += minInDay;
        this.arrInt = aInt >= minInWeek ? aInt - minInWeek : aInt;
        let fM = toMin(this.freeStr);
        let fInt = dayStart + fM;
        if (fM < aM) fInt = aInt + (fM + minInDay - aM);
        else if (aInt > (dayStart + minInDay)) fInt += minInDay;
        this.freeInt = fInt >= minInWeek ? fInt - minInWeek : fInt;
    }
}

function formatTimeStr(val) {
    if (typeof val === 'number') {
        const t = Math.round(val * 1440);
        return `${Math.floor(t/60).toString().padStart(2,'0')}:${(t%60).toString().padStart(2,'0')}`;
    }
    return String(val || "").substring(0,5);
}

// Рендер машин (з фільтрацією)
function renderVehicles() {
    const container = document.getElementById('vehicle_list');
    
    if (window.allVehicles.length === 0) {
        container.innerHTML = '<div class="empty-msg">Завантажте .xlsx з машинами</div>';
        return;
    }

    if (window.currentFilteredVehicles.length === 0) {
        container.innerHTML = '<div class="empty-msg">Машин за цими критеріями не знайдено</div>';
        return;
    }

    container.innerHTML = window.currentFilteredVehicles.map(v => {
        // Перевіряємо, чи це Шасі BDF, і призначаємо додатковий клас
        const isBDF = v.type === 'Шасі BDF';
        const badgeClass = isBDF ? 'badge badge-bdf' : 'badge';

        return `
        <div class="vehicle-card" draggable="true" ondragstart="onDragVehicle(event, '${v.number}')">
            <div style="display:flex; justify-content:space-between;">
                <strong>${v.number}</strong>
                <span class="${badgeClass}">${v.type}</span>
            </div>
            <div class="vehicle-info">
                📍 <span>${v.node}</span> | 🕒 <span>${formatMinToWeekTime(v.availableTime)}</span>
            </div>
            <button class="tp-btn" onclick="openTeleport('${v.number}')">🚀 Телепорт</button>
        </div>
        `;
    }).join('');
}

function formatMinToWeekTime(totalMin) {
    const days = ['Пн','Вт','Ср','Чт','Пт','Сб','Нд'];
    const d = Math.floor((totalMin % 10080) / 1440);
    const h = Math.floor((totalMin % 1440) / 60);
    const m = totalMin % 60;
    return `${days[d]} ${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}`;
}

// Drag & Drop
function onDragVehicle(e, vNumber) { 
    e.dataTransfer.setData('text/plain', vNumber); 
}

function allowDrop(e, row) { 
    e.preventDefault(); 
    if (row && !row.classList.contains('drag-highlight')) {
        row.classList.add('drag-highlight'); // Включаем подсветку
    }
}

function onDragLeave(e, row) {
    // Проверяем, что мы действительно покинули строку, а не просто перевели мышь на текст внутри неё
    if (row && !row.contains(e.relatedTarget)) {
        row.classList.remove('drag-highlight');
    }
}



async function onDropOnTrip(e, tripId, row) {
    e.preventDefault();
    if (row) row.classList.remove('drag-highlight');
    
    const vNumber = e.dataTransfer.getData('text/plain');
    assignCar(vNumber, tripId);
}

function assignCar(vNumber, tripId) {
    const car = window.allVehicles.find(v => v.number === vNumber);
    const trip = window.allTrips.find(t => t.id === tripId);

    if (car.type !== trip.auto) return alert(`❌ Тип ТЗ не підходить! (Треба: ${trip.auto}, є: ${car.type})`);
    if (car.node !== trip.origin) return alert(`❌ Машина на ${car.node}, а треба на ${trip.origin}`);
    if (car.availableTime > trip.trueStart) return alert("❌ Машина ще не звільниться до моменту подачі!");

    trip.assignedCar = car.number;
    car.node = trip.destination;
    car.availableTime = trip.freeInt;

    applyFilters();
}

// Телепорт
function openTeleport(vNumber) {
    currentTeleportCar = vNumber;
    document.getElementById('tp_car_name').innerText = `Машина: ${vNumber}`;
    document.getElementById('teleport_modal').style.display = 'flex';
}
function confirmTeleport() {
    const targetNode = document.getElementById('tp_node').value.trim();
    const travelTime = parseInt(document.getElementById('tp_time').value);
    
    if (!targetNode || isNaN(travelTime)) return alert("Введіть коректні дані!");

    const car = window.allVehicles.find(v => v.number === currentTeleportCar);
    
    // 1. Створюємо "фейковий" масив даних, ніби ми прочитали його з Excel
    let r = new Array(60).fill("");
    r[2] = "TELEPORT"; // GRF
    r[7] = "Перегін"; // Тип рейсу
    r[8] = car.type; // Тип авто
    r[57] = `${car.node}_${targetNode}`; // Маршрут
    
    // Визначаємо, в який день тижня відбувається телепорт (від 0 до 6)
    let dayIdx = Math.floor((car.availableTime % 10080) / 1440);
    r[16 + dayIdx] = '+'; // Астрономічний день
    r[23 + dayIdx] = '+'; // Логістичний день
    
    // 2. Створюємо об'єкт рейсу
    let tpTrip = new Trip(r);
    tpTrip.id = `tp_${Date.now()}`;
    
    // 3. Жорстко прописуємо таймінг (старт = зараз, фініш = зараз + час у дорозі)
    let startTime = car.availableTime;
    let endTime = car.availableTime + travelTime;
    
    tpTrip.origin = car.node;
    tpTrip.destination = targetNode;
    
    tpTrip.podachaInt = startTime;
    tpTrip.depInt = startTime;
    tpTrip.trueStart = startTime;
    
    tpTrip.arrInt = endTime;
    tpTrip.freeInt = endTime;
    tpTrip.trueEnd = endTime;
    
    // Форматуємо час для красивого відображення в таблиці (гг:хх)
    const formatTime = (mins) => {
        const h = Math.floor((mins % 1440) / 60);
        const m = mins % 60;
        return `${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}`;
    };
    
    tpTrip.podachaStr = formatTime(startTime);
    tpTrip.depStr = formatTime(startTime);
    tpTrip.arrStr = formatTime(endTime);
    tpTrip.freeStr = formatTime(endTime);
    
    // 4. Призначаємо машину на цей перегін
    tpTrip.assignedCar = car.number;
    
    // 5. Додаємо телепорт у загальний реєстр графіків
    window.allTrips.push(tpTrip);
    
    // 6. Оновлюємо поточний стан самої машини
    car.node = targetNode;
    car.availableTime = endTime;
    
    closeModal();
    
    // 7. Пересортовуємо таблицю, щоб телепорт встав на своє хронологічне місце
    window.allTrips.sort((a, b) => {
        if (a.logisticDay !== b.logisticDay) return a.logisticDay - b.logisticDay;
        return a.depInt - b.depInt;
    });
    
    applyFilters();
}
function closeModal() { document.getElementById('teleport_modal').style.display = 'none'; }

// Рендер таблицы
function render(trips) {
    const rows = trips.map(t => {
        const isAssigned = t.assignedCar !== null;
        return `
        <tr ondragover="allowDrop(event, this)" ondragleave="onDragLeave(event, this)" ondrop="onDropOnTrip(event, '${t.id}', this)" class="${isAssigned ? 'row-assigned' : ''}">
            <td class="car-cell">${t.assignedCar || '---'}</td>
            <td>${t.grf}</td>
            <td>${t.type || '---'}</td> 
            <td>${t.auto || '---'}</td> 
            <td title="${t.route}">${t.route}</td> 
            <td>${t.origin}</td>
            <td>${t.destination}</td>
            ${t.logDays.map(d => `<td class="col-day">${d === '+' ? '+' : ''}</td>`).join('')}
            
            <td class="time-cell">${t.podachaStr}</td>
            <td class="time-cell">${t.depStr}</td>
            <td class="time-cell">${t.arrStr}</td>
            <td class="time-cell">${t.freeStr}</td>
            
            <td>
                ${!isAssigned ? `<button onclick="quickAssign('${t.id}')" class="mini-btn">🚗</button>` : `<button onclick="unassign('${t.id}')" class="mini-btn">✕</button>`}
            </td>
        </tr>`;
    });

    if (!clusterizeInstance) {
        clusterizeInstance = new Clusterize({ rows, scrollId: 'scrollArea', contentId: 'table_body' });
    } else {
        clusterizeInstance.update(rows);
    }
}

// Кнопки дій
function quickAssign(tripId) {
    const trip = window.allTrips.find(t => t.id === tripId);
    
    // Шукаємо ТІЛЬКИ серед відфільтрованих машин (currentFilteredVehicles)
    const candidates = window.currentFilteredVehicles.filter(v => 
        v.type === trip.auto && 
        v.node === trip.origin && 
        v.availableTime <= trip.trueStart // Важливо: перевіряємо по часу подачі!
    );

    if (candidates.length === 0) {
        alert("Немає вільних машин на цьому вузлі з таким типом ТЗ (серед відфільтрованого списку).");
        return;
    }

    assignCar(candidates[0].number, tripId);
}

// 2. Скасування призначення (з урахуванням ланцюжка та видаленням телепортів)
function unassign(tripId) {
    const targetTrip = window.allTrips.find(t => t.id === tripId);
    if (!targetTrip || !targetTrip.assignedCar) return;

    const car = window.allVehicles.find(v => v.number === targetTrip.assignedCar);
    if (!car) {
        targetTrip.assignedCar = null;
        applyFilters();
        return;
    }

    // 1. Знаходимо всі рейси цієї машини і сортуємо їх хронологічно
    const carTrips = window.allTrips
        .filter(t => t.assignedCar === car.number)
        .sort((a, b) => a.trueStart - b.trueStart);

    // 2. Знаходимо індекс рейсу, який ми хочемо скасувати
    const targetIndex = carTrips.findIndex(t => t.id === tripId);

    if (targetIndex !== -1) {
        // 3. Відміняємо цей рейс і ВСІ НАСТУПНІ за ним у ланцюжку
        for (let i = targetIndex; i < carTrips.length; i++) {
            if (carTrips[i].grf === 'TELEPORT') {
                // Якщо це був створений нами телепорт — просто видаляємо його з пам'яті назавжди
                window.allTrips = window.allTrips.filter(t => t.id !== carTrips[i].id);
            } else {
                // Звичайний графік просто відкріпляємо від машини
                carTrips[i].assignedCar = null;
            }
        }

        // 4. Відновлюємо стан машини до того, який був ПЕРЕД скасованим рейсом
        if (targetIndex > 0) {
            const prevTrip = carTrips[targetIndex - 1];
            car.node = prevTrip.destination;
            car.availableTime = prevTrip.freeInt;
        } else {
            // Якщо це був самий перший рейс, скидаємо до початкового стану з файлу
            car.node = car.originalNode;
            car.availableTime = car.originalTime;
        }
    }

    applyFilters();
}

async function runAutoAssignment() {
    let count = 0;
    
    // Сортуємо і беремо ТІЛЬКИ ВІДФІЛЬТРОВАНІ графіки
    const sortedTrips = [...window.currentFilteredTrips].sort((a, b) => a.depInt - b.depInt);

    for (const trip of sortedTrips) {
        if (trip.assignedCar) continue; // Якщо вже є авто, пропускаємо

        let bestCar = null;
        let minWait = Infinity;

        // Шукаємо ТІЛЬКИ серед відфільтрованих машин
        for (const car of window.currentFilteredVehicles) {
            if (car.type === trip.auto && car.node === trip.origin && car.availableTime <= trip.trueStart) {
                let wait = trip.trueStart - car.availableTime;
                if (wait < minWait) {
                    minWait = wait;
                    bestCar = car;
                }
            }
        }

        if (bestCar) {
            trip.assignedCar = bestCar.number;
            bestCar.node = trip.destination;
            bestCar.availableTime = trip.freeInt;
            count++;
        }
    }

    applyFilters();
    alert(`Автоматично призначено ${count} рейсів.`);
}

function balanceFleet() {
    if (window.allTrips.length === 0) return alert("Спочатку завантажте графіки!");

    let required = {}; // Потреба { "Kv-ESTc": { "Шасі BDF": 5 } }
    let available = {}; // Наявність { "Kv-ESTc": { "Шасі BDF": 3 } }

    // 1. Рахуємо потребу: всі виїзди в Понеділок
    window.allTrips.forEach(t => {
        if (t.logDays[0] === '+') { // Якщо є виїзд у Пн
            if (!required[t.origin]) required[t.origin] = {};
            if (!required[t.origin][t.auto]) required[t.origin][t.auto] = 0;
            required[t.origin][t.auto]++;
        }
    });

    // 2. Рахуємо наявні авто, які доступні в Понеділок (час звільнення < 1440 хв)
    window.allVehicles.forEach(v => {
        if (v.availableTime < 1440) {
            if (!available[v.node]) available[v.node] = {};
            if (!available[v.node][v.type]) available[v.node][v.type] = 0;
            available[v.node][v.type]++;
        }
    });

    // 3. Порівнюємо і генеруємо віртуальні авто
    let addedCount = 0;

    for (const node in required) {
        for (const autoType in required[node]) {
            const reqCount = required[node][autoType];
            const avCount = (available[node] && available[node][autoType]) ? available[node][autoType] : 0;
            
            const deficit = reqCount - avCount;
            
            if (deficit > 0) {
                // Створюємо відсутні авто
                for (let i = 0; i < deficit; i++) {
                    addedCount++;
                    const prefix = node.substring(0, 3).toUpperCase();
                    window.allVehicles.push({
                        number: `VIRT-${prefix}-${Math.floor(100 + Math.random() * 900)}`, // Напр. VIRT-KV--452
                        type: autoType,
                        availableTime: 0, // Пн 00:00
                        node: node,
                        city: node,
                        originalNode: node,
                        originalTime: 0
                    });
                }
            }
        }
    }

    if (addedCount > 0) {
        alert(`Збалансовано! Додано ${addedCount} віртуальних авто (префікс VIRT-) для покриття дефіциту в понеділок.`);
        applyFilters(); // Оновлюємо списки на екрані
    } else {
        alert("Парк вже збалансований, дефіциту на понеділок немає!");
    }
}

// 5. Вивантаження результатів у Excel
function exportToExcel() {
    if (window.allTrips.length === 0 && window.allVehicles.length === 0) {
        return alert("Немає даних для вивантаження!");
    }

    // --- АРКУШ 1: МАШИНИ ---
    // Формуємо заголовки
    const carsData = [
        ["Номер ТЗ", "Тип ТЗ", "Поточний вузол", "Час звільнення"]
    ];
    
    // Додаємо дані кожної машини з її останнім станом
    window.allVehicles.forEach(v => {
        carsData.push([
            v.number, 
            v.type, 
            v.node, 
            formatMinToWeekTime(v.availableTime)
        ]);
    });

    // --- АРКУШ 2: ГРАФІКИ ---
    // Формуємо заголовки (тільки ті колонки, що є на екрані)
    const tripsData = [
        ["Авто", "GRF", "Тип", "Тип авто", "Маршрут", "Відправник", "Отримувач", 
         "Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Нд", 
         "Подача", "Виїзд", "Приїзд", "Вільний"]
    ];

    // Додаємо дані графіків
    window.allTrips.forEach(t => {
        tripsData.push([
            t.assignedCar || "", // Залишаємо пустим, якщо авто не призначено
            t.grf,
            t.type || "",
            t.auto || "",
            t.route,
            t.origin,
            t.destination,
            t.logDays[0] === '+' ? '+' : '',
            t.logDays[1] === '+' ? '+' : '',
            t.logDays[2] === '+' ? '+' : '',
            t.logDays[3] === '+' ? '+' : '',
            t.logDays[4] === '+' ? '+' : '',
            t.logDays[5] === '+' ? '+' : '',
            t.logDays[6] === '+' ? '+' : '',
            t.podachaStr,
            t.depStr,
            t.arrStr,
            t.freeStr
        ]);
    });

    // --- СТВОРЕННЯ ТА ЗБЕРЕЖЕННЯ ФАЙЛУ ---
    const wb = XLSX.utils.book_new();
    
    // Перетворюємо масиви в аркуші
    const wsCars = XLSX.utils.aoa_to_sheet(carsData);
    const wsTrips = XLSX.utils.aoa_to_sheet(tripsData);

    // Трохи магії для автоширини колонок (базовий варіант)
    wsCars['!cols'] = [{wch: 15}, {wch: 12}, {wch: 20}, {wch: 15}];
    wsTrips['!cols'] = [{wch: 12}, {wch: 12}, {wch: 20}, {wch: 15}, {wch: 35}, {wch: 15}, {wch: 15}, 
                        {wch: 4}, {wch: 4}, {wch: 4}, {wch: 4}, {wch: 4}, {wch: 4}, {wch: 4}, 
                        {wch: 8}, {wch: 8}, {wch: 8}, {wch: 8}];

    // Додаємо аркуші до книги
    XLSX.utils.book_append_sheet(wb, wsCars, "Авто");
    XLSX.utils.book_append_sheet(wb, wsTrips, "Графіки");

    // Генеруємо дату для назви файлу
    const dateStr = new Date().toISOString().slice(0, 10);
    
    // Запускаємо скачування
    XLSX.writeFile(wb, `Dispatcher_Result_${dateStr}.xlsx`);
}