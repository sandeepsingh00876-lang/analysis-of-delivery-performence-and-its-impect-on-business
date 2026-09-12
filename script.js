const files = {
    products: "avg delivery time to delivery product.csv",
    sellers: "best seller performance.csv",
    reviews: "review on delay in delivery.csv",
    routes: "state wise delivery delay.csv"
};
const colors = { ink: "#172a3a", muted: "#6c7d88", teal: "#087f8c", coral: "#ec765d", gold: "#e4a93a", grid: "#e8efeb" };
const charts = {};
const number = value => new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(value);
const days = value => `${Number(value).toFixed(1)} days`;
const clean = value => String(value ?? "").replace(/\s+/g, " ").trim();
const options = { responsive: true, maintainAspectRatio: false, animation: { duration: 700 }, plugins: { legend: { labels: { usePointStyle: true, color: colors.ink, font: { family: "DM Sans" } } }, tooltip: { backgroundColor: colors.ink, padding: 12, titleFont: { family: "Space Grotesk" }, bodyFont: { family: "DM Sans" } } }, scales: { x: { grid: { display: false }, ticks: { color: colors.muted, font: { family: "DM Sans" } } }, y: { beginAtZero: true, grid: { color: colors.grid }, ticks: { color: colors.muted, font: { family: "DM Sans" } } } } };

async function loadCsv(file) {
    const response = await fetch(encodeURI(file));
    if (!response.ok) throw new Error(`Unable to load ${file}`);
    return XLSX.utils.sheet_to_json(XLSX.read(await response.arrayBuffer(), { type: "array" }).Sheets.Sheet1);
}

function draw(id, config) {
    if (charts[id]) charts[id].destroy();
    charts[id] = new Chart(document.getElementById(id), config);
}

function renderProducts(data) {
    const rows = data.map(row => ({ category: clean(row.product_category), orders: Number(row.total_orders), delivery: Number(row.avg_delivery_days) })).filter(row => row.category && row.orders > 0).sort((a, b) => b.orders - a.orders).slice(0, 15);
    const labels = rows.map(row => row.category.replace(/_/g, " "));
    draw("productChart", {
        data: {
            labels,
            datasets: [
                { type: "bar", label: "Orders", data: rows.map(row => row.orders), backgroundColor: colors.teal, borderRadius: 4, yAxisID: "orders" },
                { type: "line", label: "Avg delivery days", data: rows.map(row => row.delivery), borderColor: colors.coral, backgroundColor: colors.coral, pointRadius: 4, tension: .3, yAxisID: "delivery" }
            ]
        },
        options: {
            ...options,
            interaction: { mode: "index", intersect: false },
            scales: {
                ...options.scales,
                orders: { ...options.scales.y, title: { display: true, text: "Orders" } },
                delivery: { ...options.scales.y, position: "right", grid: { drawOnChartArea: false }, title: { display: true, text: "Days" } },
                x: { ...options.scales.x, ticks: { ...options.scales.x.ticks, maxRotation: 55, minRotation: 35 } }
            }
        }
    });
    const total = rows.reduce((sum, row) => sum + row.orders, 0);
    const weighted = rows.reduce((sum, row) => sum + row.orders * row.delivery, 0) / total;
    document.getElementById("totalOrders").textContent = number(total);
    document.getElementById("avgDelivery").textContent = days(weighted);
    document.getElementById("topProduct").textContent = labels[0];
    document.getElementById("topProductNote").textContent = `${number(rows[0].orders)} orders`;
}

function renderSellers(data) {
    const rows = data.map(row => ({ id: clean(row.seller_id).slice(0, 8), orders: Number(row.total_orders), delivery: Number(row.avg_delivery_days), delay: Math.abs(Number(row.avg_delay_days)) })).filter(row => row.id && row.orders > 0);
    const top = [...rows].sort((a, b) => b.orders - a.orders).slice(0, 12).reverse();
    draw("sellerChart", {
        type: "bar",
        data: {
            labels: top.map(row => row.id),
            datasets: [{ label: "Orders", data: top.map(row => row.orders), backgroundColor: top.map(row => row.delay > 0 ? colors.coral : colors.teal), borderRadius: 4 }]
        },
        options: {
            ...options,
            indexAxis: "y",
            plugins: { ...options.plugins, tooltip: { ...options.plugins.tooltip, callbacks: { afterLabel: context => `Avg delay: ${days(top[context.dataIndex].delay)}` } } }
        }
    });
    const buckets = [
        { label: "0-2 days", min: 0, max: 2 },
        { label: "2-4 days", min: 2, max: 4 },
        { label: "4-6 days", min: 4, max: 6 },
        { label: "6-8 days", min: 6, max: 8 },
        { label: "8-10 days", min: 8, max: 10 },
        { label: "10+ days", min: 10, max: Infinity }
    ];
    draw("delayChart", { type: "doughnut", data: { labels: buckets.map(bucket => bucket.label), datasets: [{ data: buckets.map(bucket => rows.filter(row => row.delay >= bucket.min && row.delay < bucket.max).length), backgroundColor: ["#2a9d8f", "#73b85c", "#e9c46a", "#f4a261", "#e87852", "#c44536"], borderColor: "#ffffff", borderWidth: 2, hoverOffset: 5 }] }, options: { ...options, cutout: "64%", plugins: { ...options.plugins, legend: { ...options.plugins.legend, position: "bottom" } } } });
    document.getElementById("sellerInsight").textContent = `Delay is shown as an absolute day difference because the source values are negative. Green marks the shortest delays, yellow and orange show the middle ranges, and red marks 10+ days.`;
}

function renderReviews(data) {
    const grouped = new Map();
    data.forEach(row => { const delivery = Number(row.delivery_days); const score = Number(row.review_score); if (!Number.isFinite(delivery) || !Number.isFinite(score)) return; const current = grouped.get(delivery) || { score: 0, count: 0 }; current.score += score; current.count += 1; grouped.set(delivery, current); });
    const rows = [...grouped.entries()].sort((a, b) => a[0] - b[0]).map(([delivery, value]) => ({ delivery, score: value.score / value.count, count: value.count }));
    draw("reviewChart", { data: { labels: rows.map(row => row.delivery), datasets: [{ type: "bar", label: "Review count", data: rows.map(row => row.count), backgroundColor: "rgba(8,127,140,.14)", yAxisID: "volume" }, { type: "line", label: "Average review score", data: rows.map(row => row.score), borderColor: colors.coral, backgroundColor: colors.coral, pointRadius: 2, tension: .25, yAxisID: "score" }] }, options: { ...options, interaction: { mode: "index", intersect: false }, scales: { x: { ...options.scales.x, title: { display: true, text: "Delivery days" } }, volume: { ...options.scales.y, title: { display: true, text: "Reviews" } }, score: { ...options.scales.y, position: "right", min: 1, max: 5, grid: { drawOnChartArea: false }, title: { display: true, text: "Average score" } } } } });
    const short = rows.filter(row => row.delivery <= 10).reduce((sum, row) => sum + row.score * row.count, 0) / rows.filter(row => row.delivery <= 10).reduce((sum, row) => sum + row.count, 0);
    const long = rows.filter(row => row.delivery > 20).reduce((sum, row) => sum + row.score * row.count, 0) / rows.filter(row => row.delivery > 20).reduce((sum, row) => sum + row.count, 0);
    document.getElementById("reviewInsight").textContent = `Average satisfaction is ${short.toFixed(2)} for deliveries up to 10 days versus ${long.toFixed(2)} after 20 days. The score gap shows where delivery speed becomes a customer experience issue.`;
}

function renderRoutes(data) {
    const rows = data.map(row => ({ route: `${clean(row.seller_state)} to ${clean(row.customer_state)}`, delay: Number(row.avg_delay_days), orders: Number(row.total_orders) })).filter(row => row.route !== " to " && Number.isFinite(row.delay)).sort((a, b) => b.delay - a.delay).slice(0, 15).reverse();
    draw("routeChart", {
        type: "bar",
        data: {
            labels: rows.map(row => row.route),
            datasets: [{ label: "Average delay (days)", data: rows.map(row => row.delay), backgroundColor: rows.map(row => row.delay >= 10 ? colors.coral : colors.gold), borderRadius: 4 }]
        },
        options: {
            ...options,
            indexAxis: "y",
            plugins: { ...options.plugins, tooltip: { ...options.plugins.tooltip, callbacks: { afterLabel: context => `Orders: ${number(rows[context.dataIndex].orders)}` } } }
        }
    });
    const top = rows[rows.length - 1];
    document.getElementById("topRoute").textContent = top.route;
    document.getElementById("topRouteNote").textContent = `${days(top.delay)} average delay`;
}

async function init() {
    try {
        const [products, sellers, reviews, routes] = await Promise.all(Object.values(files).map(loadCsv));
        renderProducts(products); renderSellers(sellers); renderReviews(reviews); renderRoutes(routes);
        document.getElementById("status").textContent = `Live analysis / ${number(reviews.length)} review records processed`;
    } catch (error) {
        document.getElementById("status").textContent = "Could not load one or more CSV files. Open this page through a local web server.";
        console.error(error);
    }
}

init();

