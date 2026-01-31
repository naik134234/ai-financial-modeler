"use client";

import { useState, useEffect } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
    FileSpreadsheet,
    Building2,
    TrendingUp,
    Zap,
    Download,
    Search,
    ChevronRight,
    AlertCircle,
    CheckCircle,
    Loader2,
    BarChart3,
    DollarSign,
    LineChart,
    Sparkles,
    Database,
    Upload,
    Filter,
    ChevronDown,
    X,
    PieChart,
    Table,
} from "lucide-react";

// Types
interface Stock {
    symbol: string;
    name: string;
    sector: string;
    sector_code?: string;
}

interface CompanyInfo {
    symbol: string;
    name: string;
    sector: string;
    industry: string;
    market_cap: number | null;
    current_price: number | null;
}

interface JobStatus {
    job_id: string;
    status: string;
    progress: number;
    message: string;
    company_name?: string;
    industry?: string;
    download_url?: string;
    validation?: {
        is_valid: boolean;
        errors: any[];
    };
}

// API Functions
const API_BASE = "";

async function fetchAllStocks(sector?: string): Promise<{ stocks: Stock[]; count: number }> {
    const url = sector ? `${API_BASE}/api/stocks?sector=${sector}` : `${API_BASE}/api/stocks`;
    const response = await fetch(url);
    if (!response.ok) throw new Error("Failed to fetch stocks");
    return response.json();
}

async function searchStocks(query: string): Promise<{ results: Stock[] }> {
    const response = await fetch(`${API_BASE}/api/stocks/search/${query}`);
    if (!response.ok) throw new Error("Search failed");
    return response.json();
}

async function fetchSectors(): Promise<{ sectors: string[] }> {
    const response = await fetch(`${API_BASE}/api/sectors`);
    if (!response.ok) throw new Error("Failed to fetch sectors");
    return response.json();
}

async function generateModel(symbol: string, forecastYears: number): Promise<{ job_id: string }> {
    const response = await fetch(`${API_BASE}/api/model/generate`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            symbol,
            exchange: "NSE",
            forecast_years: forecastYears,
            model_types: ["three_statement", "dcf"],
        }),
    });
    if (!response.ok) throw new Error("Failed to start model generation");
    return response.json();
}

async function generateRawModel(data: RawModelData): Promise<{ job_id: string }> {
    const response = await fetch(`${API_BASE}/api/model/generate-raw`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
    });
    if (!response.ok) throw new Error("Failed to start raw model generation");
    return response.json();
}

async function getJobStatus(jobId: string): Promise<JobStatus> {
    const response = await fetch(`${API_BASE}/api/job/${jobId}`);
    if (!response.ok) throw new Error("Failed to get job status");
    return response.json();
}

interface RawModelData {
    company_name: string;
    industry: string;
    forecast_years: number;
    historical_data: {
        income_statement?: { revenue?: number; ebitda?: number; net_income?: number };
        balance_sheet?: { total_assets?: number; total_liabilities?: number };
    };
    assumptions: {
        revenue_growth?: number;
        ebitda_margin?: number;
        tax_rate?: number;
    };
}

// Feature cards data
const FEATURES = [
    {
        icon: Zap,
        title: "AI-Powered",
        description: "Gemini AI classifies industry & builds model logic",
        color: "from-yellow-500 to-orange-500",
    },
    {
        icon: FileSpreadsheet,
        title: "Real Excel Formulas",
        description: "Linked cells, charts, sensitivity analysis",
        color: "from-green-500 to-emerald-500",
    },
    {
        icon: Building2,
        title: "150+ Stocks",
        description: "Power, Banking, FMCG, IT, Pharma & more",
        color: "from-blue-500 to-cyan-500",
    },
    {
        icon: TrendingUp,
        title: "Advanced DCF",
        description: "WACC, scenarios, dashboard with charts",
        color: "from-purple-500 to-pink-500",
    },
];

const SECTOR_COLORS: { [key: string]: string } = {
    power: "bg-yellow-500/20 text-yellow-400 border-yellow-500/30",
    banking: "bg-blue-500/20 text-blue-400 border-blue-500/30",
    it: "bg-purple-500/20 text-purple-400 border-purple-500/30",
    pharma: "bg-green-500/20 text-green-400 border-green-500/30",
    fmcg: "bg-orange-500/20 text-orange-400 border-orange-500/30",
    auto: "bg-red-500/20 text-red-400 border-red-500/30",
    metals: "bg-slate-500/20 text-slate-400 border-slate-500/30",
    oil_gas: "bg-amber-500/20 text-amber-400 border-amber-500/30",
    cement: "bg-stone-500/20 text-stone-400 border-stone-500/30",
    infra: "bg-cyan-500/20 text-cyan-400 border-cyan-500/30",
    nbfc: "bg-indigo-500/20 text-indigo-400 border-indigo-500/30",
    telecom: "bg-pink-500/20 text-pink-400 border-pink-500/30",
    chemicals: "bg-lime-500/20 text-lime-400 border-lime-500/30",
    consumer: "bg-rose-500/20 text-rose-400 border-rose-500/30",
};

type TabType = "stocks" | "raw";

export default function Home() {
    const [activeTab, setActiveTab] = useState<TabType>("stocks");
    const [stocks, setStocks] = useState<Stock[]>([]);
    const [sectors, setSectors] = useState<string[]>([]);
    const [selectedSector, setSelectedSector] = useState<string>("");
    const [searchQuery, setSearchQuery] = useState("");
    const [searchResults, setSearchResults] = useState<Stock[]>([]);

    const [selectedStock, setSelectedStock] = useState<Stock | null>(null);
    const [forecastYears, setForecastYears] = useState(5);
    const [jobStatus, setJobStatus] = useState<JobStatus | null>(null);
    const [error, setError] = useState<string | null>(null);
    const [isLoading, setIsLoading] = useState(false);

    // Raw data form
    const [rawData, setRawData] = useState<RawModelData>({
        company_name: "",
        industry: "general",
        forecast_years: 5,
        historical_data: {
            income_statement: { revenue: 10000, ebitda: 2500, net_income: 1500 },
            balance_sheet: { total_assets: 20000, total_liabilities: 8000 },
        },
        assumptions: {
            revenue_growth: 0.10,
            ebitda_margin: 0.25,
            tax_rate: 0.25,
        },
    });

    // Load stocks and sectors
    useEffect(() => {
        fetchSectors().then((data) => setSectors(data.sectors)).catch(console.error);
        fetchAllStocks().then((data) => setStocks(data.stocks)).catch(console.error);
    }, []);

    // Filter by sector
    useEffect(() => {
        if (selectedSector) {
            fetchAllStocks(selectedSector).then((data) => setStocks(data.stocks)).catch(console.error);
        } else {
            fetchAllStocks().then((data) => setStocks(data.stocks)).catch(console.error);
        }
    }, [selectedSector]);

    // Search stocks
    useEffect(() => {
        if (searchQuery.length >= 2) {
            searchStocks(searchQuery).then((data) => setSearchResults(data.results)).catch(console.error);
        } else {
            setSearchResults([]);
        }
    }, [searchQuery]);

    // Generate model
    const handleGenerate = async () => {
        if (activeTab === "stocks" && !selectedStock) return;
        if (activeTab === "raw" && !rawData.company_name.trim()) return;

        setIsLoading(true);
        setError(null);
        setJobStatus(null);

        try {
            let job_id: string;

            if (activeTab === "stocks" && selectedStock) {
                const result = await generateModel(selectedStock.symbol, forecastYears);
                job_id = result.job_id;
            } else {
                const result = await generateRawModel(rawData);
                job_id = result.job_id;
            }

            setJobStatus({ job_id, status: "pending", progress: 0, message: "Starting..." });
        } catch (err) {
            setError("Failed to start model generation. Please try again.");
            setIsLoading(false);
        }
    };

    // Poll job status
    useEffect(() => {
        if (!jobStatus || jobStatus.status === "completed" || jobStatus.status === "failed") {
            setIsLoading(false);
            return;
        }

        const interval = setInterval(async () => {
            try {
                const status = await getJobStatus(jobStatus.job_id);
                setJobStatus(status);

                if (status.status === "completed" || status.status === "failed") {
                    setIsLoading(false);
                    clearInterval(interval);
                }
            } catch (err) {
                console.error("Failed to get job status:", err);
            }
        }, 1000);

        return () => clearInterval(interval);
    }, [jobStatus?.job_id, jobStatus?.status]);

    const displayedStocks = searchQuery.length >= 2 ? searchResults : stocks;

    return (
        <main className="min-h-screen py-8 px-4 sm:px-6 lg:px-8">
            <div className="max-w-7xl mx-auto">
                {/* Header */}
                <motion.div
                    initial={{ opacity: 0, y: -20 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="text-center mb-8"
                >
                    <div className="flex items-center justify-center gap-3 mb-4">
                        <div className="p-3 rounded-2xl bg-gradient-to-br from-primary-500 to-purple-600 shadow-glow-md">
                            <BarChart3 className="w-8 h-8 text-white" />
                        </div>
                        <h1 className="text-4xl sm:text-5xl font-bold bg-gradient-to-r from-white via-primary-200 to-purple-200 bg-clip-text text-transparent">
                            AI Financial Modeler
                        </h1>
                    </div>
                    <p className="text-dark-300 text-lg max-w-2xl mx-auto">
                        Generate institutional-grade Excel models with{" "}
                        <span className="text-primary-400">DCF, charts, sensitivity analysis</span>
                        {" "}for 150+ Indian stocks
                    </p>
                </motion.div>

                {/* Features Grid */}
                <motion.div
                    initial={{ opacity: 0, y: 20 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: 0.1 }}
                    className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-8"
                >
                    {FEATURES.map((feature, index) => (
                        <motion.div
                            key={feature.title}
                            initial={{ opacity: 0, y: 20 }}
                            animate={{ opacity: 1, y: 0 }}
                            transition={{ delay: 0.1 + index * 0.05 }}
                            className="card p-4 group hover:border-primary-500/30 transition-all duration-300"
                        >
                            <div className={`w-10 h-10 rounded-xl bg-gradient-to-br ${feature.color} p-2 mb-3 group-hover:scale-110 transition-transform`}>
                                <feature.icon className="w-full h-full text-white" />
                            </div>
                            <h3 className="font-semibold text-white mb-1">{feature.title}</h3>
                            <p className="text-sm text-dark-400">{feature.description}</p>
                        </motion.div>
                    ))}
                </motion.div>

                {/* Main Content */}
                <div className="grid lg:grid-cols-3 gap-6">
                    {/* Left Panel - Stock Selection */}
                    <motion.div
                        initial={{ opacity: 0, x: -20 }}
                        animate={{ opacity: 1, x: 0 }}
                        transition={{ delay: 0.2 }}
                        className="lg:col-span-2 card p-6"
                    >
                        {/* Tab Navigation */}
                        <div className="flex gap-2 mb-6">
                            <button
                                onClick={() => setActiveTab("stocks")}
                                className={`flex items-center gap-2 px-4 py-2 rounded-xl font-medium transition-all ${activeTab === "stocks"
                                        ? "bg-primary-500 text-white"
                                        : "bg-dark-700/50 text-dark-300 hover:bg-dark-600"
                                    }`}
                            >
                                <Database className="w-4 h-4" />
                                Select Stock
                            </button>
                            <button
                                onClick={() => setActiveTab("raw")}
                                className={`flex items-center gap-2 px-4 py-2 rounded-xl font-medium transition-all ${activeTab === "raw"
                                        ? "bg-primary-500 text-white"
                                        : "bg-dark-700/50 text-dark-300 hover:bg-dark-600"
                                    }`}
                            >
                                <Upload className="w-4 h-4" />
                                Raw Data Input
                            </button>
                        </div>

                        {activeTab === "stocks" ? (
                            <>
                                {/* Search & Filter */}
                                <div className="flex flex-wrap gap-3 mb-4">
                                    <div className="relative flex-1 min-w-[200px]">
                                        <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-5 h-5 text-dark-400" />
                                        <input
                                            type="text"
                                            value={searchQuery}
                                            onChange={(e) => setSearchQuery(e.target.value.toUpperCase())}
                                            placeholder="Search stocks..."
                                            className="input-field w-full pl-12"
                                        />
                                    </div>
                                    <div className="relative">
                                        <select
                                            value={selectedSector}
                                            onChange={(e) => setSelectedSector(e.target.value)}
                                            className="input-field pr-10 appearance-none"
                                        >
                                            <option value="">All Sectors</option>
                                            {sectors.map((sector) => (
                                                <option key={sector} value={sector}>
                                                    {sector.replace("_", " ").toUpperCase()}
                                                </option>
                                            ))}
                                        </select>
                                        <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-5 h-5 text-dark-400 pointer-events-none" />
                                    </div>
                                </div>

                                {/* Stocks Grid */}
                                <div className="max-h-[400px] overflow-y-auto custom-scrollbar">
                                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                        {displayedStocks.map((stock) => (
                                            <button
                                                key={stock.symbol}
                                                onClick={() => {
                                                    setSelectedStock(stock);
                                                    setJobStatus(null);
                                                    setError(null);
                                                }}
                                                className={`p-3 rounded-xl border text-left transition-all ${selectedStock?.symbol === stock.symbol
                                                        ? "border-primary-500 bg-primary-500/10"
                                                        : "border-dark-600 hover:border-dark-500 bg-dark-800/50"
                                                    }`}
                                            >
                                                <div className="flex items-center justify-between">
                                                    <div>
                                                        <p className="font-semibold text-white">{stock.symbol}</p>
                                                        <p className="text-sm text-dark-400 truncate max-w-[180px]">{stock.name}</p>
                                                    </div>
                                                    <span className={`text-xs px-2 py-1 rounded-lg border ${SECTOR_COLORS[stock.sector_code || ""] || "bg-dark-600/50 text-dark-300 border-dark-500"
                                                        }`}>
                                                        {stock.sector_code?.replace("_", " ")}
                                                    </span>
                                                </div>
                                            </button>
                                        ))}
                                    </div>
                                </div>
                            </>
                        ) : (
                            /* Raw Data Input Form */
                            <div className="space-y-4">
                                <div className="grid md:grid-cols-2 gap-4">
                                    <div>
                                        <label className="block text-sm font-medium text-dark-300 mb-2">
                                            Company Name *
                                        </label>
                                        <input
                                            type="text"
                                            value={rawData.company_name}
                                            onChange={(e) => setRawData({ ...rawData, company_name: e.target.value })}
                                            placeholder="e.g., My Company Ltd"
                                            className="input-field w-full"
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-dark-300 mb-2">
                                            Industry
                                        </label>
                                        <select
                                            value={rawData.industry}
                                            onChange={(e) => setRawData({ ...rawData, industry: e.target.value })}
                                            className="input-field w-full"
                                        >
                                            <option value="general">General Corporate</option>
                                            <option value="power">Power & Utilities</option>
                                            <option value="banking">Banking</option>
                                            <option value="it">IT Services</option>
                                            <option value="pharma">Pharmaceuticals</option>
                                            <option value="fmcg">FMCG</option>
                                            <option value="auto">Automobiles</option>
                                            <option value="metals">Metals & Mining</option>
                                            <option value="oil_gas">Oil & Gas</option>
                                            <option value="cement">Cement</option>
                                            <option value="infra">Infrastructure</option>
                                        </select>
                                    </div>
                                </div>

                                {/* Historical Data */}
                                <div className="p-4 rounded-xl bg-dark-800/50 border border-dark-600">
                                    <h4 className="font-medium text-white mb-3 flex items-center gap-2">
                                        <Table className="w-4 h-4 text-primary-400" />
                                        Historical Data (₹ Crores)
                                    </h4>
                                    <div className="grid md:grid-cols-3 gap-3">
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Revenue</label>
                                            <input
                                                type="number"
                                                value={rawData.historical_data.income_statement?.revenue || 0}
                                                onChange={(e) => setRawData({
                                                    ...rawData,
                                                    historical_data: {
                                                        ...rawData.historical_data,
                                                        income_statement: {
                                                            ...rawData.historical_data.income_statement,
                                                            revenue: parseFloat(e.target.value) || 0,
                                                        },
                                                    },
                                                })}
                                                className="input-field w-full text-sm"
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">EBITDA</label>
                                            <input
                                                type="number"
                                                value={rawData.historical_data.income_statement?.ebitda || 0}
                                                onChange={(e) => setRawData({
                                                    ...rawData,
                                                    historical_data: {
                                                        ...rawData.historical_data,
                                                        income_statement: {
                                                            ...rawData.historical_data.income_statement,
                                                            ebitda: parseFloat(e.target.value) || 0,
                                                        },
                                                    },
                                                })}
                                                className="input-field w-full text-sm"
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Net Income</label>
                                            <input
                                                type="number"
                                                value={rawData.historical_data.income_statement?.net_income || 0}
                                                onChange={(e) => setRawData({
                                                    ...rawData,
                                                    historical_data: {
                                                        ...rawData.historical_data,
                                                        income_statement: {
                                                            ...rawData.historical_data.income_statement,
                                                            net_income: parseFloat(e.target.value) || 0,
                                                        },
                                                    },
                                                })}
                                                className="input-field w-full text-sm"
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Total Assets</label>
                                            <input
                                                type="number"
                                                value={rawData.historical_data.balance_sheet?.total_assets || 0}
                                                onChange={(e) => setRawData({
                                                    ...rawData,
                                                    historical_data: {
                                                        ...rawData.historical_data,
                                                        balance_sheet: {
                                                            ...rawData.historical_data.balance_sheet,
                                                            total_assets: parseFloat(e.target.value) || 0,
                                                        },
                                                    },
                                                })}
                                                className="input-field w-full text-sm"
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Total Liabilities</label>
                                            <input
                                                type="number"
                                                value={rawData.historical_data.balance_sheet?.total_liabilities || 0}
                                                onChange={(e) => setRawData({
                                                    ...rawData,
                                                    historical_data: {
                                                        ...rawData.historical_data,
                                                        balance_sheet: {
                                                            ...rawData.historical_data.balance_sheet,
                                                            total_liabilities: parseFloat(e.target.value) || 0,
                                                        },
                                                    },
                                                })}
                                                className="input-field w-full text-sm"
                                            />
                                        </div>
                                    </div>
                                </div>

                                {/* Assumptions */}
                                <div className="p-4 rounded-xl bg-dark-800/50 border border-dark-600">
                                    <h4 className="font-medium text-white mb-3 flex items-center gap-2">
                                        <PieChart className="w-4 h-4 text-green-400" />
                                        Model Assumptions
                                    </h4>
                                    <div className="grid md:grid-cols-3 gap-3">
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Revenue Growth %</label>
                                            <input
                                                type="number"
                                                step="0.01"
                                                value={(rawData.assumptions.revenue_growth || 0) * 100}
                                                onChange={(e) => setRawData({
                                                    ...rawData,
                                                    assumptions: {
                                                        ...rawData.assumptions,
                                                        revenue_growth: (parseFloat(e.target.value) || 0) / 100,
                                                    },
                                                })}
                                                className="input-field w-full text-sm"
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">EBITDA Margin %</label>
                                            <input
                                                type="number"
                                                step="0.01"
                                                value={(rawData.assumptions.ebitda_margin || 0) * 100}
                                                onChange={(e) => setRawData({
                                                    ...rawData,
                                                    assumptions: {
                                                        ...rawData.assumptions,
                                                        ebitda_margin: (parseFloat(e.target.value) || 0) / 100,
                                                    },
                                                })}
                                                className="input-field w-full text-sm"
                                            />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-dark-400 mb-1">Tax Rate %</label>
                                            <input
                                                type="number"
                                                step="0.01"
                                                value={(rawData.assumptions.tax_rate || 0) * 100}
                                                onChange={(e) => setRawData({
                                                    ...rawData,
                                                    assumptions: {
                                                        ...rawData.assumptions,
                                                        tax_rate: (parseFloat(e.target.value) || 0) / 100,
                                                    },
                                                })}
                                                className="input-field w-full text-sm"
                                            />
                                        </div>
                                    </div>
                                </div>
                            </div>
                        )}

                        {/* Error Display */}
                        <AnimatePresence>
                            {error && (
                                <motion.div
                                    initial={{ opacity: 0, height: 0 }}
                                    animate={{ opacity: 1, height: "auto" }}
                                    exit={{ opacity: 0, height: 0 }}
                                    className="mt-4"
                                >
                                    <div className="flex items-center gap-3 p-4 rounded-xl bg-red-500/10 border border-red-500/30 text-red-400">
                                        <AlertCircle className="w-5 h-5 flex-shrink-0" />
                                        <p>{error}</p>
                                    </div>
                                </motion.div>
                            )}
                        </AnimatePresence>
                    </motion.div>

                    {/* Right Panel - Configuration & Progress */}
                    <motion.div
                        initial={{ opacity: 0, x: 20 }}
                        animate={{ opacity: 1, x: 0 }}
                        transition={{ delay: 0.3 }}
                        className="card p-6"
                    >
                        {/* Selected Stock / Company Info */}
                        {(selectedStock || (activeTab === "raw" && rawData.company_name)) && (
                            <div className="mb-6">
                                <div className="p-4 rounded-xl bg-gradient-to-r from-primary-500/10 to-purple-500/10 border border-primary-500/20">
                                    <div className="flex items-center gap-2 mb-1">
                                        <Building2 className="w-5 h-5 text-primary-400" />
                                        <h3 className="font-semibold text-white">
                                            {activeTab === "stocks" ? selectedStock?.name : rawData.company_name}
                                        </h3>
                                    </div>
                                    <p className="text-sm text-dark-400">
                                        {activeTab === "stocks"
                                            ? `${selectedStock?.symbol} • ${selectedStock?.sector}`
                                            : `Custom Data • ${rawData.industry.replace("_", " ")}`
                                        }
                                    </p>
                                </div>
                            </div>
                        )}

                        {/* Configuration */}
                        {!jobStatus && (
                            <div className="space-y-4">
                                <div>
                                    <label className="block text-sm font-medium text-dark-300 mb-2">
                                        Forecast Period
                                    </label>
                                    <select
                                        value={forecastYears}
                                        onChange={(e) => {
                                            setForecastYears(parseInt(e.target.value));
                                            if (activeTab === "raw") {
                                                setRawData({ ...rawData, forecast_years: parseInt(e.target.value) });
                                            }
                                        }}
                                        className="input-field w-full"
                                    >
                                        {[3, 5, 7, 10].map((years) => (
                                            <option key={years} value={years}>
                                                {years} Years
                                            </option>
                                        ))}
                                    </select>
                                </div>

                                <button
                                    onClick={handleGenerate}
                                    disabled={
                                        isLoading ||
                                        (activeTab === "stocks" && !selectedStock) ||
                                        (activeTab === "raw" && !rawData.company_name.trim())
                                    }
                                    className="btn-primary w-full flex items-center justify-center gap-2"
                                >
                                    {isLoading ? (
                                        <Loader2 className="w-5 h-5 animate-spin" />
                                    ) : (
                                        <Sparkles className="w-5 h-5" />
                                    )}
                                    Generate Model
                                    <ChevronRight className="w-5 h-5" />
                                </button>

                                {/* Model Features */}
                                <div className="pt-4 border-t border-dark-600">
                                    <h4 className="text-sm font-medium text-dark-300 mb-3">Model Includes:</h4>
                                    <ul className="space-y-2">
                                        {[
                                            "Income Statement (5Y Historical + Forecast)",
                                            "Balance Sheet with Balance Check",
                                            "Cash Flow Statement",
                                            "DCF Valuation with WACC",
                                            "Sensitivity Analysis",
                                            "Scenario Analysis (Bear/Base/Bull)",
                                            "Dashboard with Charts",
                                        ].map((item) => (
                                            <li key={item} className="flex items-center gap-2 text-sm text-dark-400">
                                                <CheckCircle className="w-4 h-4 text-green-400" />
                                                {item}
                                            </li>
                                        ))}
                                    </ul>
                                </div>
                            </div>
                        )}

                        {/* Progress Display */}
                        <AnimatePresence>
                            {jobStatus && (
                                <motion.div
                                    initial={{ opacity: 0, y: 10 }}
                                    animate={{ opacity: 1, y: 0 }}
                                >
                                    <div className="space-y-4">
                                        <div className="flex items-center justify-between">
                                            <div className="flex items-center gap-3">
                                                {jobStatus.status === "completed" ? (
                                                    <CheckCircle className="w-6 h-6 text-green-400" />
                                                ) : jobStatus.status === "failed" ? (
                                                    <AlertCircle className="w-6 h-6 text-red-400" />
                                                ) : (
                                                    <Loader2 className="w-6 h-6 text-primary-400 animate-spin" />
                                                )}
                                                <div>
                                                    <p className="font-medium text-white">
                                                        {jobStatus.company_name || selectedStock?.name || rawData.company_name}
                                                    </p>
                                                    <p className="text-sm text-dark-400">{jobStatus.industry}</p>
                                                </div>
                                            </div>
                                            <span className={`px-3 py-1 rounded-full text-sm font-medium ${jobStatus.status === "completed"
                                                    ? "bg-green-500/20 text-green-400"
                                                    : jobStatus.status === "failed"
                                                        ? "bg-red-500/20 text-red-400"
                                                        : "bg-primary-500/20 text-primary-400"
                                                }`}>
                                                {jobStatus.status === "completed" ? "Ready" :
                                                    jobStatus.status === "failed" ? "Failed" :
                                                        `${jobStatus.progress}%`}
                                            </span>
                                        </div>

                                        {/* Progress bar */}
                                        <div className="h-2 bg-dark-600 rounded-full overflow-hidden">
                                            <motion.div
                                                initial={{ width: 0 }}
                                                animate={{ width: `${jobStatus.progress}%` }}
                                                className="h-full progress-bar rounded-full"
                                            />
                                        </div>

                                        <p className="text-sm text-dark-400">{jobStatus.message}</p>

                                        {/* Download button */}
                                        {jobStatus.status === "completed" && jobStatus.download_url && (
                                            <motion.div
                                                initial={{ opacity: 0, y: 10 }}
                                                animate={{ opacity: 1, y: 0 }}
                                            >
                                                <a
                                                    href={jobStatus.download_url}
                                                    className="btn-primary w-full inline-flex items-center justify-center gap-2"
                                                >
                                                    <Download className="w-5 h-5" />
                                                    Download Excel Model
                                                </a>

                                                {jobStatus.validation && (
                                                    <div className="mt-4 p-4 rounded-lg bg-dark-800/50">
                                                        <div className="flex items-center gap-2 mb-2">
                                                            {jobStatus.validation.is_valid ? (
                                                                <CheckCircle className="w-4 h-4 text-green-400" />
                                                            ) : (
                                                                <AlertCircle className="w-4 h-4 text-yellow-400" />
                                                            )}
                                                            <span className="text-sm font-medium text-white">
                                                                Model Validation
                                                            </span>
                                                        </div>
                                                        {jobStatus.validation.errors.length > 0 && (
                                                            <ul className="text-sm text-dark-400 space-y-1">
                                                                {jobStatus.validation.errors.slice(0, 3).map((err, i) => (
                                                                    <li key={i}>• {err.message}</li>
                                                                ))}
                                                            </ul>
                                                        )}
                                                    </div>
                                                )}
                                            </motion.div>
                                        )}

                                        {/* New model button */}
                                        {(jobStatus.status === "completed" || jobStatus.status === "failed") && (
                                            <button
                                                onClick={() => {
                                                    setJobStatus(null);
                                                    setSelectedStock(null);
                                                    setRawData({
                                                        ...rawData,
                                                        company_name: "",
                                                    });
                                                }}
                                                className="btn-secondary w-full mt-2"
                                            >
                                                Generate Another Model
                                            </button>
                                        )}
                                    </div>
                                </motion.div>
                            )}
                        </AnimatePresence>
                    </motion.div>
                </div>

                {/* Footer */}
                <motion.footer
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    transition={{ delay: 0.5 }}
                    className="text-center text-dark-500 text-sm mt-8"
                >
                    <p className="flex items-center justify-center gap-2">
                        <LineChart className="w-4 h-4" />
                        Built with AI • DCF Valuation • Charts & Sensitivity Analysis • 150+ Stocks
                    </p>
                </motion.footer>
            </div>
        </main>
    );
}
