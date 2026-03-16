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

function addListenerIfExist(id, event, fn) {
    const el = document.getElementById(id);
    if (el) el.addEventListener(event, fn);
}

// Слушаем новые и старые ID
addListenerIfExist('node_filter_trips', 'input', applyFilters); // Если переименовал верхний фильтр
addListenerIfExist('node_filter', 'input', applyFilters);       // Если оставил старое название верхнего фильтра
addListenerIfExist('node_filter_cars', 'input', applyFilters);  // НАШ НОВЫЙ ФИЛЬТР АВТО
addListenerIfExist('filter_car_number', 'input', applyFilters);
// ДОДАЛИ СЛУХАЧ ДЛЯ ТИПУ АВТО:
addListenerIfExist('filter_car_type', 'input', applyFilters);

// Прив'язка подій до всіх фільтрів для миттєвої реакції
document.getElementById('node_filter_trips').addEventListener('input', applyFilters);
document.getElementById('node_filter_cars').addEventListener('input', applyFilters);
document.getElementById('filter_car_number').addEventListener('input', applyFilters);
document.querySelectorAll('.trip-filter').forEach(el => el.addEventListener('input', applyFilters));
document.querySelectorAll('.trip-filter').forEach(el => el.addEventListener('change', applyFilters));


// Кастомний Alert (тільки кнопка ОК)
window.customAlert = function(message, title = "ℹ️ Інформація") {
    return new Promise(resolve => {
        document.getElementById('dialog_title').innerText = title;
        document.getElementById('dialog_message').innerText = message;
        document.getElementById('btn_dialog_cancel').style.display = 'none'; // Ховаємо кнопку скасування
        document.getElementById('custom_dialog_modal').style.display = 'flex';

        document.getElementById('btn_dialog_ok').onclick = () => {
            document.getElementById('custom_dialog_modal').style.display = 'none';
            resolve(true);
        };
    });
};

// Кастомний Confirm (кнопки ОК та Скасувати)
window.customConfirm = function(message, title = "❓ Підтвердження") {
    return new Promise(resolve => {
        document.getElementById('dialog_title').innerText = title;
        document.getElementById('dialog_message').innerText = message;
        document.getElementById('btn_dialog_cancel').style.display = 'block'; // Показуємо скасування
        document.getElementById('custom_dialog_modal').style.display = 'flex';

        document.getElementById('btn_dialog_ok').onclick = () => {
            document.getElementById('custom_dialog_modal').style.display = 'none';
            resolve(true);
        };
        document.getElementById('btn_dialog_cancel').onclick = () => {
            document.getElementById('custom_dialog_modal').style.display = 'none';
            resolve(false);
        };
    });
};

// Універсальна функція застосування фільтрів
function applyFilters() {
    // Зчитуємо вузли графіків
    const elTripNode = document.getElementById('node_filter_trips') || document.getElementById('node_filter');
    const tripNodeFilter = elTripNode ? elTripNode.value : "";
    
    // Зчитуємо фільтри МАШИН
    const elCarNode = document.getElementById('node_filter_cars');
    const elCarNum = document.getElementById('filter_car_number');
    const elCarType = document.getElementById('filter_car_type'); 

    const carNodeFilter = elCarNode ? elCarNode.value : "";
    const carNumberFilter = elCarNum ? elCarNum.value : "";
    const carTypeFilter = elCarType ? elCarType.value : ""; 
    
    // Зчитуємо інші фільтри графіків
    const dayFilter = document.getElementById('filter_day') ? document.getElementById('filter_day').value : "";
    const statusFilter = document.getElementById('filter_status') ? document.getElementById('filter_status').value : "";
    const autoTypeFilter = document.getElementById('filter_auto_type') ? document.getElementById('filter_auto_type').value : "";
    const tripTypeFilter = document.getElementById('filter_trip_type') ? document.getElementById('filter_trip_type').value : "";
    const assignedCarFilter = document.getElementById('filter_assigned_car') ? document.getElementById('filter_assigned_car').value : "";
    const destFilter = document.getElementById('filter_destination') ? document.getElementById('filter_destination').value : "";

    // 1. Фільтруємо графіки (використовуємо multiMatch)
    window.currentFilteredTrips = window.allTrips.filter(t => {
        if (!multiMatch(t.origin, tripNodeFilter)) return false;
        if (!multiMatch(t.destination, destFilter)) return false;
        
        // Випадаючі списки залишаємо як є
        if (dayFilter !== "" && t.logDays[parseInt(dayFilter)] !== '+') return false;
        if (statusFilter === 'with_auto' && !t.assignedCar) return false;
        if (statusFilter === 'without_auto' && t.assignedCar) return false;
        
        if (!multiMatch(t.auto, autoTypeFilter)) return false;
        if (!multiMatch(t.type, tripTypeFilter)) return false;
        if (!multiMatch(t.assignedCar, assignedCarFilter)) return false;
        
        return true; 
    });

    // 2. Фільтруємо машини (використовуємо multiMatch)
    window.currentFilteredVehicles = window.allVehicles.filter(v => {
        if (!multiMatch(v.node, carNodeFilter)) return false;
        if (!multiMatch(v.number, carNumberFilter)) return false;
        if (!multiMatch(v.type, carTypeFilter)) return false;
        return true;
    });

    // Відмальовуємо
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
            const initTime = parseExcelTime(r[2]); // Залишаємо математику для розрахунків
            
            // Робимо точне відображення дати без прив'язки до поточного тижня
            let rawText = "---";
            if (typeof r[2] === 'number') {
                if (r[2] < 1) {
                    // Якщо вказано тільки час (наприклад, 0.5 = 12:00)
                    rawText = formatMinToWeekTime(initTime);
                } else {
                    // Якщо вказана повноцінна дата, витягуємо її напряму
                    // Excel рахує дні з 1900 року, 25569 - це різниця до 1970 року (Unix Epoch)
                    const jsDate = new Date(Math.round((r[2] - 25569) * 86400 * 1000));
                    const days = ['Нд', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
                    
                    const dd = String(jsDate.getUTCDate()).padStart(2, '0');
                    const mm = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
                    const dayStr = days[jsDate.getUTCDay()];
                    const hh = String(jsDate.getUTCHours()).padStart(2, '0');
                    const min = String(jsDate.getUTCMinutes()).padStart(2, '0');
                    
                    rawText = `${dd}.${mm} (${dayStr}) ${hh}:${min}`;
                }
            } else if (r[2] !== undefined) {
                rawText = String(r[2]);
            }

            return {
                number: String(r[0]),
                type: String(r[1]),
                availableTime: initTime, 
                displayTime: rawText, 
                node: String(r[3]).trim(),
                city: r[5],
                originalNode: String(r[3]).trim(),
                originalTime: initTime, 
                originalDisplayTime: rawText 
            };
        });
        applyFilters();
    }
}

function parseExcelTime(val) {
    if (typeof val === 'number') {
        if (val < 1) {
            return Math.round(val * 1440);
        } else {
            // Дістаємо дату машини з Excel
            let jsDateUTC = new Date(Math.round((val - 25569) * 86400 * 1000));
            let localExcelDate = new Date(
                jsDateUTC.getUTCFullYear(),
                jsDateUTC.getUTCMonth(),
                jsDateUTC.getUTCDate(),
                jsDateUTC.getUTCHours(),
                jsDateUTC.getUTCMinutes()
            );

            // Беремо стартовий понеділок З НАШОГО ПОЛЯ
            let baseMonday = getBaseMonday();
            
            // Рахуємо різницю в хвилинах
            let timeDiffMins = Math.floor((localExcelDate - baseMonday) / 60000);
            
            // Якщо машина звільнилася ДО початку планування (наприклад, у неділю), 
            // для графіка вона готова з нульової хвилини понеділка
            if (timeDiffMins <= 0) {
                return 0; 
            }
            
            return timeDiffMins;
        }
    }

    if (typeof val === 'string') {
        const days = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Нд'];
        let d = 0;
        days.forEach((day, idx) => { if (val.includes(day)) d = idx; });
        
        const timeMatch = val.match(/(\d{1,2}):(\d{2})/);
        if (timeMatch) {
            return d * 1440 + parseInt(timeMatch[1]) * 60 + parseInt(timeMatch[2]);
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
        const minInDay = 1440;
        
        // Базовый астрономический день
        let actualAstroDay = this.dayIndex;
        
        // Корректируем смещение недель:
        // Если логистический Вс (6), а астро Пн (0) -> разница -6. Это Пн следующей недели!
        if (actualAstroDay - this.logisticDay < -3) {
            actualAstroDay += 7;
        } 
        // Если логистический Пн (0), а астро Вс (6) -> разница +6. Это Вс предыдущей недели!
        else if (actualAstroDay - this.logisticDay > 3) {
            actualAstroDay -= 7;
        }

        const dayStart = actualAstroDay * minInDay; // Теперь стартуем с правильных суток

        const toMin = (str) => {
            if (!str) return 0;
            const [h, m] = str.split(':').map(Number);
            return (h * 60) + m;
        };

        const pM = toMin(this.podachaStr);
        const dM = toMin(this.depStr);
        const aM = toMin(this.arrStr);
        const fM = toMin(this.freeStr);

        // 1. Точка отсчета - выезд (departure)
        this.depInt = dayStart + dM;

        // 2. Подача (podacha) - происходит ДО выезда. Считаем время ожидания/погрузки
        let loadTime = dM - pM;
        if (loadTime < 0) loadTime += minInDay; // Если погрузка началась до полуночи
        this.podachaInt = this.depInt - loadTime;

        // 3. Приезд (arrival) - происходит ПОСЛЕ выезда. Считаем время в пути
        let transitTime = aM - dM;
        if (transitTime < 0) transitTime += minInDay; // Если в пути пересекли полночь
        this.arrInt = this.depInt + transitTime;

        // 4. Освобождение (free) - происходит ПОСЛЕ приезда. Считаем время разгрузки
        let unloadTime = fM - aM;
        if (unloadTime < 0) unloadTime += minInDay; // Если разгрузка перешла за полночь
        this.freeInt = this.arrInt + unloadTime;
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
                📍 <span>${v.node}</span> | 🕒 <span>${v.displayTime}</span>
            </div>
            <button class="tp-btn" onclick="openTeleport('${v.number}')">🚀 Телепорт</button>
        </div>
        `;
    }).join('');
}

// Знаходимо понеділок поточного тижня як точку відліку
const dateInput = document.getElementById('plan_start_date');

function initStartDate() {
    const now = new Date();
    const day = now.getDay();
    const diff = now.getDate() - day + (day === 0 ? -6 : 1);
    const monday = new Date(now.setDate(diff));
    
    // Форматуємо для поля <input type="date"> (РРРР-ММ-ДД)
    const yyyy = monday.getFullYear();
    const mm = String(monday.getMonth() + 1).padStart(2, '0');
    const dd = String(monday.getDate()).padStart(2, '0');
    dateInput.value = `${yyyy}-${mm}-${dd}`;
}
initStartDate();

// Функція для отримання обраного понеділка з поля
function getBaseMonday() {
    const parts = dateInput.value.split('-');
    return new Date(parts[0], parts[1] - 1, parts[2]);
}
const baseMonday = getBaseMonday();


function formatMinToWeekTime(totalMin) {
    const days = ['Нд', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб']; // Формат Date.getDay()
    
    // 1. Берем наш стартовый понедельник из поля на экране
    const startMonday = getBaseMonday();
    
    // 2. Превращаем минуты в миллисекунды и просто прибавляем к понедельнику
    const targetDate = new Date(startMonday.getTime() + totalMin * 60000);
    
    // 3. Достаем реальные данные
    const dd = String(targetDate.getDate()).padStart(2, '0');
    const mm = String(targetDate.getMonth() + 1).padStart(2, '0');
    const dayStr = days[targetDate.getDay()];
    const h = String(targetDate.getHours()).padStart(2, '0');
    const m = String(targetDate.getMinutes()).padStart(2, '0');
    
    return `${dd}.${mm} (${dayStr}) ${h}:${m}`;
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

async function assignCar(vNumber, tripId) {
    const car = window.allVehicles.find(v => v.number === vNumber);
    const trip = window.allTrips.find(t => t.id === tripId);

    const bufferBdf = parseInt(document.getElementById('buffer_bdf').value) || 0;
    const bufferOther = parseInt(document.getElementById('buffer_other').value) || 0;
    
    const isBDF = (car.type || "").toLowerCase().includes('бдф') || (car.type || "").toLowerCase().includes('bdf');
    const requiredBuffer = isBDF ? bufferBdf : bufferOther;

    if ((car.type || "").toLowerCase() !== (trip.auto || "").toLowerCase()) {
        // НОВИЙ ДОБРИЙ CONFIRM
        const confirmChange = await customConfirm(`Тип ТЗ не співпадає! \nПлан: ${trip.auto}\nФакт: ${car.type}\n\nПризначити все одно?`, "⚠️ Невідповідність типу");
        if (!confirmChange) return;
    }

    if ((car.node || "").toLowerCase() !== (trip.origin || "").toLowerCase()) {
        // НОВИЙ ДОБРИЙ ALERT
        await customAlert(`Машина на ${car.node}, а треба на ${trip.origin}`, "❌ Помилка вузла");
        return;
    }

    const maxAllowedTime = trip.depInt - requiredBuffer; 

    if (car.availableTime > maxAllowedTime) {
        await customAlert(`Машина критично спізнюється!\n\nГотовність: ${formatMinToWeekTime(car.availableTime)}\nКрайній термін (Виїзд мінус ${requiredBuffer} хв): ${formatMinToWeekTime(maxAllowedTime)}\n\nПризначення неможливе.`, "🚫 Блокування");
        return;
    } 
    else if (car.availableTime > trip.podachaInt) {
        const forceAssign = await customConfirm(`Машина запізнюється на планову подачу!\n\nГотовність: ${formatMinToWeekTime(car.availableTime)}\nПодача за графіком: ${formatMinToWeekTime(trip.podachaInt)}\nВиїзд: ${formatMinToWeekTime(trip.depInt)}\n\nВона встигає лише до крайнього терміну (${formatMinToWeekTime(maxAllowedTime)}). Призначити примусово?`, "⏳ Запізнення");
        if (!forceAssign) return;
    }

    trip.assignedCar = car.number;
    trip.assignedCarType = car.type; 
    car.node = trip.destination;

    car.availableTime = trip.arrInt + requiredBuffer;
    car.displayTime = formatMinToWeekTime(car.availableTime);

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
    car.displayTime = formatMinToWeekTime(endTime);
    
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
            <td style="color: #1a73e8; font-weight: bold;">
                ${t.assignedCar ? (t.assignedCarType || window.allVehicles.find(v => v.number === t.assignedCar)?.type || '---') : '---'}
            </td>
            <td title="${t.route}">${t.route}</td> 
            <td>${t.origin}</td>
            <td>${t.destination}</td>
            ${t.logDays.map(d => `<td class="col-day">${d === '+' ? '+' : ''}</td>`).join('')}
            
            <td class="time-cell" title="${formatMinToWeekTime(t.podachaInt)}">${t.podachaStr}</td>
            <td class="time-cell" title="${formatMinToWeekTime(t.depInt)}">${t.depStr}</td>
            <td class="time-cell" title="${formatMinToWeekTime(t.arrInt)}">${t.arrStr}</td>
            <td class="time-cell" title="${formatMinToWeekTime(t.freeInt)}">${t.freeStr}</td>
            
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
async function quickAssign(tripId) {
    const trip = window.allTrips.find(t => t.id === tripId);
    
    // Шукаємо ТІЛЬКИ серед відфільтрованих машин (currentFilteredVehicles)
    const candidates = window.currentFilteredVehicles.filter(v => 
        (v.type || "").toLowerCase() === (trip.auto || "").toLowerCase() && 
        (v.node || "").toLowerCase() === (trip.origin || "").toLowerCase() && 
        v.availableTime <= trip.trueStart // Важливо: перевіряємо по часу подачі!
    );

    if (candidates.length === 0) {
        alert("Немає вільних машин на цьому вузлі з таким типом ТЗ (серед відфільтрованого списку).");
        return;
    }

    await assignCar(candidates[0].number, tripId);
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
                window.allTrips = window.allTrips.filter(t => t.id !== carTrips[i].id);
            } else {
                carTrips[i].assignedCar = null;
                carTrips[i].assignedCarType = null; // Очищаем факт
            }
        }

        // 4. Відновлюємо стан машини до того, який був ПЕРЕД скасованим рейсом
        if (targetIndex > 0) {
            const prevTrip = carTrips[targetIndex - 1];
            car.node = prevTrip.destination;
            car.availableTime = prevTrip.freeInt;
            car.displayTime = formatMinToWeekTime(prevTrip.freeInt);
        } else {
            // Якщо це був самий перший рейс, скидаємо до початкового стану з файлу
            car.node = car.originalNode;
            car.availableTime = car.originalTime;
            car.displayTime = car.originalDisplayTime;
        }
    }

    applyFilters();
}

async function runAutoAssignment() {
    let totalCount = 0;
    let assignedInThisIteration;

    // Крутимо цикл доти, доки в ітерації була призначена хоча б одна машина
    do {
        assignedInThisIteration = 0;
        const sortedTrips = [...window.currentFilteredTrips].sort((a, b) => a.depInt - b.depInt);

        for (const trip of sortedTrips) {
            if (trip.assignedCar) continue;

            let bestCar = null;
            let minWait = Infinity;

            for (const car of window.currentFilteredVehicles) {
                // Сравниваем тип и узел без учета регистра
                if ((car.type || "").toLowerCase() === (trip.auto || "").toLowerCase() && 
                    (car.node || "").toLowerCase() === (trip.origin || "").toLowerCase() && 
                    car.availableTime <= trip.trueStart) {
                    
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
                
                // --- ДОДАЙ ОСЬ ЦЕЙ РЯДОК ---
                bestCar.displayTime = formatMinToWeekTime(trip.freeInt); 
                // ---------------------------

                assignedInThisIteration++;
                totalCount++;
            }
        }
    } while (assignedInThisIteration > 0);

    applyFilters();
    await customAlert(`Автоматично призначено ${totalCount} рейсів.`);
    //alert(`Автоматично призначено ${totalCount} рейсів.`);
}

async function runAutoAssignmentByDeparture() {
    const bufferBdf = parseInt(document.getElementById('buffer_bdf').value) || 0;
    const bufferOther = parseInt(document.getElementById('buffer_other').value) || 0;
    
    let totalCount = 0;
    let assignedInThisIteration;

    do {
        assignedInThisIteration = 0;
        const sortedTrips = [...window.currentFilteredTrips].sort((a, b) => a.depInt - b.depInt);

        for (const trip of sortedTrips) {
            if (trip.assignedCar) continue;

            let bestCar = null;
            let minWait = Infinity;

            for (const car of window.currentFilteredVehicles) {
                if ((car.type || "").toLowerCase() === (trip.auto || "").toLowerCase() && 
                    (car.node || "").toLowerCase() === (trip.origin || "").toLowerCase()) {
                    
                    // Визначаємо необхідний буфер залежно від типу авто
                    const isBDF = (car.type || "").toLowerCase().includes('бдф') || (car.type || "").toLowerCase().includes('bdf');
                    const requiredBuffer = isBDF ? bufferBdf : bufferOther;
                    
                    // УМОВА 1: Готовність машини + хвилини <= Час виїзду
                    if (car.availableTime + requiredBuffer <= trip.depInt) {
                        let wait = trip.depInt - (car.availableTime + requiredBuffer);
                        if (wait < minWait) {
                            minWait = wait;
                            bestCar = car;
                        }
                    }
                }
            }

            if (bestCar) {
                // Призначаємо машину
                trip.assignedCar = bestCar.number;
                trip.assignedCarType = bestCar.type; // Фіксуємо факт типу
                bestCar.node = trip.destination;
                
                // Знову перевіряємо тип обраної машини для буферу після рейсу
                const isBDF = (bestCar.type || "").toLowerCase().includes('бдф') || (bestCar.type || "").toLowerCase().includes('bdf');
                const requiredBuffer = isBDF ? bufferBdf : bufferOther;

                // УМОВА 2: Нова готовність = Приїзд + хвилини на розвантаження
                bestCar.availableTime = trip.arrInt + requiredBuffer;
                bestCar.displayTime = formatMinToWeekTime(bestCar.availableTime); 
                
                assignedInThisIteration++;
                totalCount++;
            }
        }
    } while (assignedInThisIteration > 0);

    applyFilters();
    //alert(`Автоматично призначено ${totalCount} рейсів!\n(Новий час готовності розраховано за схемою: Приїзд + буфер)`);
    await customAlert(`Автоматично призначено ${totalCount} рейсів!\n(Новий час готовності розраховано за схемою: Приїзд + буфер)`);
}

// Нова функція: повне скасування всіх призначень
async function cancelAllAssignments() {
    if (!(await customConfirm("Ви впевнені, що хочете скасувати ВСІ призначення та видалити створені телепорти?", "Скидання"))) return;

    // 1. Очищаємо графіки
    window.allTrips.forEach(t => {
        t.assignedCar = null;
    });

    // 2. Видаляємо всі "фейкові" рейси-телепорти, якщо вони були створені
    window.allTrips = window.allTrips.filter(t => t.grf !== 'TELEPORT');

    // 3. Повертаємо машини на їхні початкові позиції
    window.allVehicles.forEach(v => {
        v.node = v.originalNode;
        v.availableTime = v.originalTime;
        v.displayTime = v.originalDisplayTime;
    });

    applyFilters();
    await customAlert("Всі призначення успішно скасовано!", "✅ Готово");
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
        ["Авто", "GRF", "Тип", "Тип авто", "Тип авто (факт)", "Маршрут", "Відправник", "Отримувач", 
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
            t.assignedCar ? t.assignedCarType : "",
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

function multiMatch(targetValue, filterString) {
    if (!filterString) return true; // Якщо фільтр пустий, пропускаємо (все ок)
    
    const target = String(targetValue || "").toLowerCase();
    
    // Розбиваємо рядок фільтра по комі, прибираємо зайві пробіли і порожні шматки
    const parts = filterString.split(',')
                              .map(s => s.trim().toLowerCase())
                              .filter(s => s.length > 0);
    
    if (parts.length === 0) return true;
    
    // Перевіряємо, чи є хоча б один збіг
    return parts.some(part => target.includes(part));
}