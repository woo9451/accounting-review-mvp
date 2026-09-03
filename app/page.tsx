"use client";

import {
  AlertTriangle,
  ArrowDownToLine,
  BarChart3,
  CheckCircle2,
  ClipboardCheck,
  Database,
  FileSpreadsheet,
  History,
  MessageSquareText,
  Search,
  ShieldAlert,
  SlidersHorizontal,
  Upload,
  XCircle,
} from "lucide-react";
import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type ReviewStatus = "검토 후보" | "소명 요청" | "답변 기록" | "수정 요청" | "반려" | "완료";

type Transaction = {
  id: string;
  date: string;
  department: string;
  vendor: string;
  category: string;
  description: string;
  amount: number;
  budget: number;
  evidenceAmount?: number;
  status: ReviewStatus;
  answer?: string;
};

type Finding = {
  id: string;
  transaction: Transaction;
  types: string[];
  severity: number;
  reason: string;
  request: string;
};

declare global {
  interface Document {
    modelContext?: {
      registerTool: (
        tool: {
          name: string;
          title?: string;
          description: string;
          inputSchema: object;
          annotations?: { readOnlyHint?: boolean; untrustedContentHint?: boolean };
          execute: (input: unknown) => unknown;
        },
        options?: { signal?: AbortSignal },
      ) => void | Promise<void>;
    };
  }
}

const sampleRows: Transaction[] = [
  {
    id: "TX-240901-001",
    date: "2026-09-01",
    department: "마케팅팀",
    vendor: "브라이트미디어",
    category: "광고선전비",
    description: "하반기 캠페인 집행",
    amount: 4_260_000,
    budget: 3_800_000,
    evidenceAmount: 4_260_000,
    status: "검토 후보",
  },
  {
    id: "TX-240901-002",
    date: "2026-09-01",
    department: "영업1팀",
    vendor: "스카이호텔",
    category: "접대비",
    description: "거래처 미팅",
    amount: 3_400_000,
    budget: 5_000_000,
    evidenceAmount: 3_100_000,
    status: "소명 요청",
    answer: "영수증 재발급 요청 중",
  },
  {
    id: "TX-240901-003",
    date: "2026-09-01",
    department: "영업1팀",
    vendor: "스카이호텔",
    category: "접대비",
    description: "거래처 미팅",
    amount: 3_400_000,
    budget: 5_000_000,
    evidenceAmount: 3_100_000,
    status: "답변 기록",
    answer: "동일 건 분할 입력 가능성 확인 필요",
  },
  {
    id: "TX-240901-004",
    date: "2026-09-01",
    department: "개발팀",
    vendor: "클라우드랩",
    category: "소프트웨어",
    description: "서버 사용료",
    amount: 0,
    budget: 2_000_000,
    evidenceAmount: 88_000,
    status: "수정 요청",
  },
  {
    id: "TX-240902-001",
    date: "2026-09-02",
    department: "인사팀",
    vendor: "한빛교육원",
    category: "교육훈련비",
    description: "신입 온보딩 교육",
    amount: 1_280_000,
    budget: 2_300_000,
    evidenceAmount: 1_280_000,
    status: "완료",
  },
  {
    id: "TX-240902-002",
    date: "2026-09-02",
    department: "총무팀",
    vendor: "오피스플러스",
    category: "소모품비",
    description: "분기 사무용품",
    amount: -52_000,
    budget: 900_000,
    evidenceAmount: 52_000,
    status: "반려",
  },
];

const statusOptions: ReviewStatus[] = [
  "검토 후보",
  "소명 요청",
  "답변 기록",
  "수정 요청",
  "반려",
  "완료",
];

const columnAliases: Record<keyof Omit<Transaction, "status" | "answer">, string[]> = {
  id: ["id", "거래id", "거래번호", "전표번호", "문서번호"],
  date: ["date", "거래일자", "일자", "사용일자", "작성일"],
  department: ["department", "부서", "사용부서", "귀속부서", "코스트센터"],
  vendor: ["vendor", "거래처", "가맹점", "공급자", "상호"],
  category: ["category", "계정과목", "예산항목", "항목", "비목"],
  description: ["description", "내용", "적요", "품목", "메모"],
  amount: ["amount", "금액", "지출금액", "공급가액", "승인금액"],
  budget: ["budget", "예산", "예산금액", "잔여예산", "승인예산"],
  evidenceAmount: ["evidenceamount", "증빙금액", "영수증금액", "세금계산서금액"],
};

function normalizeHeader(value: string) {
  return value.toLowerCase().replace(/[\s_\-()]/g, "");
}

function parseNumber(value: string | number | undefined) {
  if (typeof value === "number") return value;
  if (!value) return 0;
  const parsed = Number(String(value).replace(/[^0-9.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatMoney(value: number) {
  return new Intl.NumberFormat("ko-KR").format(value);
}

function parseCsv(text: string) {
  const rows: string[][] = [];
  let current = "";
  let row: string[] = [];
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"' && next === '"') {
      current += '"';
      index += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if ((char === "," || char === "\t") && !inQuotes) {
      row.push(current.trim());
      current = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(current.trim());
      if (row.some(Boolean)) rows.push(row);
      row = [];
      current = "";
    } else {
      current += char;
    }
  }

  row.push(current.trim());
  if (row.some(Boolean)) rows.push(row);
  return rows;
}

async function readSpreadsheet(file: File) {
  if (file.name.toLowerCase().endsWith(".xlsx") || file.name.toLowerCase().endsWith(".xls")) {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json<string[]>(firstSheet, { header: 1, defval: "" });
  }

  return parseCsv(await file.text());
}

function mapRows(rows: string[][]): Transaction[] {
  const headers = rows[0]?.map(normalizeHeader) ?? [];
  const indexes = Object.fromEntries(
    Object.keys(columnAliases).map((key) => {
      const aliases = columnAliases[key as keyof typeof columnAliases].map(normalizeHeader);
      return [key, headers.findIndex((header) => aliases.includes(header))];
    }),
  ) as Record<keyof Omit<Transaction, "status" | "answer">, number>;

  return rows.slice(1).map((row, index) => ({
    id: row[indexes.id] || `UPLOAD-${String(index + 1).padStart(3, "0")}`,
    date: row[indexes.date] || "날짜 없음",
    department: row[indexes.department] || "부서 미지정",
    vendor: row[indexes.vendor] || "거래처 미지정",
    category: row[indexes.category] || "항목 미지정",
    description: row[indexes.description] || "내용 없음",
    amount: parseNumber(row[indexes.amount]),
    budget: parseNumber(row[indexes.budget]),
    evidenceAmount:
      indexes.evidenceAmount >= 0 ? parseNumber(row[indexes.evidenceAmount]) : undefined,
    status: "검토 후보",
  }));
}

function detectFindings(rows: Transaction[]): Finding[] {
  const duplicateKeys = new Map<string, number>();
  rows.forEach((row) => {
    const key = `${row.date}|${row.department}|${row.vendor}|${row.amount}`;
    duplicateKeys.set(key, (duplicateKeys.get(key) ?? 0) + 1);
  });

  return rows
    .map((row) => {
      const duplicateKey = `${row.date}|${row.department}|${row.vendor}|${row.amount}`;
      const types: string[] = [];

      if ((duplicateKeys.get(duplicateKey) ?? 0) > 1) types.push("중복 거래");
      if (row.budget > 0 && row.amount > row.budget) types.push("부서별 예산 초과");
      if (row.amount >= 3_000_000) types.push("고액 지출");
      if (!Number.isFinite(row.amount) || row.amount <= 0) types.push("비정상 금액");
      if (typeof row.evidenceAmount === "number" && row.evidenceAmount !== row.amount) {
        types.push("증빙 금액 불일치");
      }

      if (types.length === 0) return null;

      const severity = Math.min(100, 38 + types.length * 16 + (row.amount >= 3_000_000 ? 10 : 0));
      const reason = [
        types.includes("중복 거래") &&
          `거래일자, 금액, 부서, 거래처가 같은 거래가 ${duplicateKeys.get(duplicateKey)}건 존재합니다.`,
        types.includes("부서별 예산 초과") &&
          `${row.department}의 ${row.category} 지출이 예산보다 ${formatMoney(row.amount - row.budget)}원 높습니다.`,
        types.includes("고액 지출") && "건별 금액이 300만 원 이상입니다.",
        types.includes("비정상 금액") && "금액이 숫자가 아니거나 0원 이하입니다.",
        types.includes("증빙 금액 불일치") &&
          `증빙 금액 ${formatMoney(row.evidenceAmount ?? 0)}원과 지출내역 ${formatMoney(row.amount)}원이 다릅니다.`,
      ]
        .filter(Boolean)
        .join(" ");

      return {
        id: `${row.id}-${types.join("-")}`,
        transaction: row,
        types,
        severity,
        reason,
        request: `${row.department} ${row.vendor} ${formatMoney(
          row.amount,
        )}원 지출 건에서 ${types.join(", ")} 항목이 확인되었습니다. 증빙과 예산 승인 근거를 회신해 주세요.`,
      };
    })
    .filter(Boolean) as Finding[];
}

function downloadCsv(findings: Finding[]) {
  const headers = [
    "거래ID",
    "거래일자",
    "부서",
    "거래처",
    "계정과목",
    "금액",
    "예산",
    "오류유형",
    "위험도",
    "상태",
    "탐지이유",
    "소명요청문",
  ];
  const body = findings.map(({ transaction, types, severity, reason, request }) =>
    [
      transaction.id,
      transaction.date,
      transaction.department,
      transaction.vendor,
      transaction.category,
      transaction.amount,
      transaction.budget,
      types.join(" / "),
      severity,
      transaction.status,
      reason,
      request,
    ]
      .map((cell) => `"${String(cell).replace(/"/g, '""')}"`)
      .join(","),
  );
  const csv = `\uFEFF${[headers.join(","), ...body].join("\n")}`;
  const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = "accounting-review-findings.csv";
  anchor.click();
  URL.revokeObjectURL(url);
}

export default function Home() {
  const [rows, setRows] = useState<Transaction[]>(sampleRows);
  const [query, setQuery] = useState("");
  const [activeType, setActiveType] = useState("전체");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);

  const findings = useMemo(() => detectFindings(rows), [rows]);
  const filteredFindings = findings.filter((finding) => {
    const haystack = [
      finding.transaction.id,
      finding.transaction.department,
      finding.transaction.vendor,
      finding.transaction.category,
      finding.transaction.description,
      finding.types.join(" "),
      finding.transaction.status,
    ]
      .join(" ")
      .toLowerCase();
    return haystack.includes(query.toLowerCase()) && (activeType === "전체" || finding.types.includes(activeType));
  });

  const selected = filteredFindings.find((finding) => finding.id === selectedId) ?? filteredFindings[0];
  const totalAmount = rows.reduce((sum, row) => sum + Math.max(row.amount, 0), 0);
  const overBudget = findings.filter((finding) => finding.types.includes("부서별 예산 초과")).length;
  const mismatch = findings.filter((finding) => finding.types.includes("증빙 금액 불일치")).length;
  const typeCounts = ["중복 거래", "부서별 예산 초과", "고액 지출", "비정상 금액", "증빙 금액 불일치"].map(
    (type) => ({ type, count: findings.filter((finding) => finding.types.includes(type)).length }),
  );

  const handleUpload = async (file: File | undefined) => {
    if (!file) return;
    const parsed = mapRows(await readSpreadsheet(file));
    if (parsed.length > 0) setRows(parsed);
  };

  const updateStatus = (id: string, status: ReviewStatus) => {
    setRows((current) => current.map((row) => (row.id === id ? { ...row, status } : row)));
  };

  const scrollToSection = (id: string) => {
    document.getElementById(id)?.scrollIntoView({
      behavior: "smooth",
      block: id === "summary" ? "start" : "center",
    });
    window.history.replaceState(null, "", `#${id}`);
  };

  useEffect(() => {
    const context = document.modelContext;
    if (!context?.registerTool) return;

    const lifecycle = new AbortController();
    const register = context.registerTool.bind(context);

    try {
      void Promise.resolve(
        register(
          {
            name: "read_accounting_review_summary",
            title: "회계 검토 요약 읽기",
            description: "현재 화면의 회계 이상데이터 검토 건수와 오류 유형별 집계를 읽습니다.",
            inputSchema: { type: "object", properties: {}, additionalProperties: false },
            annotations: { readOnlyHint: true, untrustedContentHint: false },
            execute: () => ({
              transactions: rows.length,
              findings: findings.length,
              byType: typeCounts,
              visibleFindings: filteredFindings.length,
            }),
          },
          { signal: lifecycle.signal },
        ),
      );

      void Promise.resolve(
        register(
          {
            name: "update_accounting_review_status",
            title: "회계 검토 상태 변경",
            description: "거래 ID나 탐지 ID를 받아 검토 상태를 변경합니다.",
            inputSchema: {
              type: "object",
              properties: {
                id: { type: "string" },
                status: { type: "string", enum: statusOptions },
              },
              required: ["id", "status"],
              additionalProperties: false,
            },
            annotations: { readOnlyHint: false, untrustedContentHint: false },
            execute: (input) => {
              const payload = input as { id?: string; status?: ReviewStatus };
              if (!payload.id || !payload.status || !statusOptions.includes(payload.status)) {
                throw new Error("id와 올바른 상태값이 필요합니다.");
              }

              const target = findings.find(
                (finding) => finding.id === payload.id || finding.transaction.id === payload.id,
              );
              if (!target) throw new Error("해당 검토 건을 찾을 수 없습니다.");

              updateStatus(target.transaction.id, payload.status);
              return { id: target.transaction.id, status: payload.status };
            },
          },
          { signal: lifecycle.signal },
        ),
      );
    } catch (error) {
      console.error(error);
    }

    return () => lifecycle.abort();
  }, [filteredFindings.length, findings, rows.length, typeCounts]);

  return (
    <main className="min-h-screen bg-background text-foreground">
      <header className="sticky top-0 z-20 border-b bg-card/95 backdrop-blur">
        <div className="mx-auto flex h-16 max-w-7xl items-center justify-between gap-4 px-5 lg:px-8">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-md bg-primary text-primary-foreground">
              <ShieldAlert className="h-5 w-5" />
            </div>
            <div>
              <p className="text-sm font-semibold leading-none">AuditFlow</p>
              <p className="mt-1 text-xs text-muted-foreground">Accounting Review MVP</p>
            </div>
          </div>

          <nav className="hidden items-center gap-1 rounded-md bg-muted p-1 text-sm font-medium text-muted-foreground md:flex">
            {[
              ["요약", "summary"],
              ["탐지 규칙", "rules"],
              ["검토함", "review"],
              ["소명 이력", "history"],
            ].map(([label, id]) => (
              <button
                key={id}
                className="rounded px-3 py-2 transition hover:bg-card hover:text-foreground"
                onClick={() => scrollToSection(id)}
                type="button"
              >
                {label}
              </button>
            ))}
          </nav>
          <div className="hidden w-[140px] lg:block" aria-hidden="true" />
        </div>
      </header>

      <section id="summary" className="scroll-mt-20 border-b bg-[linear-gradient(135deg,#f8faf7_0%,#eef6f4_50%,#fff8e8_100%)]">
        <div className="mx-auto grid max-w-7xl gap-8 px-5 py-10 lg:grid-cols-[minmax(0,1fr)_420px] lg:px-8">
          <div className="flex flex-col justify-center">
            <p className="inline-flex w-fit items-center gap-2 rounded-md border bg-card px-3 py-1 text-sm font-medium text-muted-foreground shadow-sm">
              <BarChart3 className="h-4 w-4 text-primary" />
              회계 이상데이터 검토 MVP
            </p>
            <h1 className="mt-5 max-w-3xl text-4xl font-semibold tracking-tight text-foreground lg:text-5xl">
              예산 초과와 증빙 오류를 보기 좋은 검토함으로 정리합니다
            </h1>
            <p className="mt-4 max-w-2xl text-base leading-7 text-muted-foreground">
              엑셀을 업로드하면 중복 거래, 고액 지출, 예산 초과, 비정상 금액, 증빙 금액 불일치를 자동으로 선별하고 소명 문안까지 준비합니다.
            </p>
            <div className="mt-6 flex flex-wrap gap-3">
              <input
                ref={fileRef}
                className="hidden"
                type="file"
                accept=".xlsx,.xls,.csv,.tsv,.txt"
                onChange={(event) => handleUpload(event.target.files?.[0])}
              />
              <button
                className="inline-flex h-11 items-center gap-2 rounded-md bg-primary px-5 text-sm font-semibold text-primary-foreground shadow-sm"
                onClick={() => fileRef.current?.click()}
              >
                <Upload className="h-4 w-4" />
                엑셀/CSV 업로드
              </button>
              <button
                className="inline-flex h-11 items-center gap-2 rounded-md border bg-card px-5 text-sm font-semibold shadow-sm"
                onClick={() => downloadCsv(filteredFindings)}
              >
                <ArrowDownToLine className="h-4 w-4" />
                오류 목록 다운로드
              </button>
            </div>
          </div>

          <div className="rounded-lg border bg-card p-5 shadow-sm">
            <div className="mb-4 flex items-center justify-between">
              <div>
                <p className="text-sm font-semibold">오늘의 검토 현황</p>
                <p className="mt-1 text-xs text-muted-foreground">샘플 데이터 기준 실시간 집계</p>
              </div>
              <Database className="h-5 w-5 text-primary" />
            </div>
            <div className="grid gap-3 sm:grid-cols-2">
            {[
              ["검토 대상", `${rows.length}건`, FileSpreadsheet],
              ["탐지 오류", `${findings.length}건`, ShieldAlert],
              ["예산 초과", `${overBudget}건`, AlertTriangle],
              ["증빙 불일치", `${mismatch}건`, ClipboardCheck],
            ].map(([label, value, Icon]) => (
              <div key={label as string} className="rounded-md border bg-background p-4">
                <div className="flex items-center justify-between text-muted-foreground">
                  <span className="text-sm">{label as string}</span>
                  <Icon className="h-4 w-4" />
                </div>
                <p className="mt-3 text-2xl font-semibold">{value as string}</p>
              </div>
            ))}
            </div>
          </div>
        </div>
      </section>

      <section className="mx-auto grid max-w-7xl gap-6 px-5 py-7 lg:grid-cols-[300px_minmax(0,1fr)_360px] lg:px-8">
        <aside className="space-y-5">
          <div id="rules" className="scroll-mt-24 rounded-lg border bg-card p-5 shadow-sm">
            <div className="flex items-center gap-2">
              <SlidersHorizontal className="h-4 w-4 text-primary" />
              <p className="text-sm font-semibold">열 자동 매핑</p>
            </div>
            <p className="mt-2 text-sm leading-6 text-muted-foreground">
              거래일자, 금액, 부서, 거래처, 예산, 증빙금액 열을 자동 인식합니다.
            </p>
            <div className="mt-4 space-y-2 text-sm">
              {["거래일자", "부서", "거래처", "금액", "예산", "증빙금액"].map((item) => (
                <div key={item} className="flex items-center justify-between rounded-md bg-muted px-3 py-2">
                  <span>{item}</span>
                  <CheckCircle2 className="h-4 w-4 text-emerald-600" />
                </div>
              ))}
            </div>
          </div>

          <div className="rounded-lg border bg-card p-5 shadow-sm">
            <p className="text-sm font-semibold">오류 유형별 검토함</p>
            <div className="mt-4 space-y-3">
              {[{ type: "전체", count: findings.length }, ...typeCounts].map((item) => (
                <button
                  key={item.type}
                  className={`w-full rounded-md px-3 py-3 text-left text-sm transition ${
                    activeType === item.type ? "bg-primary text-primary-foreground" : "bg-muted"
                  }`}
                  onClick={() => setActiveType(item.type)}
                >
                  <span className="flex items-center justify-between">
                    <span>{item.type}</span>
                    <span>{item.count}</span>
                  </span>
                  <span className={`mt-2 block h-1.5 rounded-full ${activeType === item.type ? "bg-white/35" : "bg-card"}`}>
                    <span
                      className={`block h-1.5 rounded-full ${activeType === item.type ? "bg-white" : "bg-primary"}`}
                      style={{ width: `${findings.length ? Math.max(8, (item.count / findings.length) * 100) : 0}%` }}
                    />
                  </span>
                </button>
              ))}
            </div>
          </div>
        </aside>

        <section id="review" className="scroll-mt-24 space-y-4">
          <div className="flex flex-col gap-3 rounded-lg border bg-card p-4 shadow-sm md:flex-row md:items-center">
            <div className="relative flex-1">
              <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
              <input
                className="h-10 w-full rounded-md border bg-background pl-9 pr-3 text-sm outline-none focus:ring-2 focus:ring-ring"
                placeholder="부서, 거래처, 오류 유형, 상태 검색"
                value={query}
                onChange={(event) => setQuery(event.target.value)}
              />
            </div>
            <div className="text-sm text-muted-foreground">
              총 지출 {formatMoney(totalAmount)}원 중 위험 거래 {filteredFindings.length}건
            </div>
          </div>

          <div className="overflow-hidden rounded-lg border bg-card shadow-sm">
            <div className="grid grid-cols-[1.1fr_.9fr_.9fr_.8fr_.8fr] border-b bg-muted px-4 py-3 text-xs font-semibold text-muted-foreground max-md:hidden">
              <span>거래</span>
              <span>유형</span>
              <span>금액/예산</span>
              <span>위험도</span>
              <span>상태</span>
            </div>
            <div className="divide-y">
              {filteredFindings.map((finding) => (
                <button
                  key={finding.id}
                  className={`grid w-full gap-3 px-4 py-4 text-left transition hover:bg-muted/70 md:grid-cols-[1.1fr_.9fr_.9fr_.8fr_.8fr] ${
                    selected?.id === finding.id ? "bg-muted" : ""
                  }`}
                  onClick={() => setSelectedId(finding.id)}
                >
                  <div>
                    <p className="font-medium">{finding.transaction.vendor}</p>
                    <p className="mt-1 text-sm text-muted-foreground">
                      {finding.transaction.date} · {finding.transaction.department} · {finding.transaction.category}
                    </p>
                  </div>
                  <div className="flex flex-wrap gap-1">
                    {finding.types.map((type) => (
                      <span key={type} className="rounded-md bg-amber-100 px-2 py-1 text-xs font-medium text-amber-900">
                        {type}
                      </span>
                    ))}
                  </div>
                  <div className="text-sm">
                    <p className="font-medium">{formatMoney(finding.transaction.amount)}원</p>
                    <p className="text-muted-foreground">예산 {formatMoney(finding.transaction.budget)}원</p>
                  </div>
                  <div>
                    <div className="h-2 w-full rounded-full bg-muted">
                      <span className="block h-2 rounded-full bg-destructive" style={{ width: `${finding.severity}%` }} />
                    </div>
                    <p className="mt-2 text-sm font-medium">{finding.severity}점</p>
                  </div>
                  <select
                    className="h-9 rounded-md border bg-background px-2 text-sm"
                    value={finding.transaction.status}
                    onChange={(event) => updateStatus(finding.transaction.id, event.target.value as ReviewStatus)}
                    onClick={(event) => event.stopPropagation()}
                  >
                    {statusOptions.map((status) => (
                      <option key={status}>{status}</option>
                    ))}
                  </select>
                </button>
              ))}
            </div>
          </div>
        </section>

        <aside id="history" className="scroll-mt-24 space-y-5">
          {selected ? (
            <>
              <div className="rounded-lg border bg-card p-5 shadow-sm">
                <div className="flex items-start justify-between gap-4">
                  <div>
                    <p className="text-sm font-semibold">탐지 이유 설명</p>
                    <h2 className="mt-2 text-xl font-semibold">{selected.transaction.vendor}</h2>
                  </div>
                  <XCircle className="h-5 w-5 text-destructive" />
                </div>
                <p className="mt-4 text-sm leading-6 text-muted-foreground">{selected.reason}</p>
              </div>

              <div className="rounded-lg border bg-card p-5 shadow-sm">
                <div className="flex items-center gap-2">
                  <MessageSquareText className="h-4 w-4" />
                  <p className="text-sm font-semibold">소명 요청 자동 문안</p>
                </div>
                <textarea
                  className="mt-3 min-h-32 w-full resize-none rounded-md border bg-background p-3 text-sm leading-6 outline-none focus:ring-2 focus:ring-ring"
                  value={selected.request}
                  readOnly
                />
              </div>

              <div className="rounded-lg border bg-card p-5 shadow-sm">
                <div className="flex items-center gap-2">
                  <History className="h-4 w-4" />
                  <p className="text-sm font-semibold">처리 이력</p>
                </div>
                <div className="mt-4 space-y-3 text-sm">
                  <p className="rounded-md bg-muted p-3">자동 탐지 완료 · 위험도 {selected.severity}점</p>
                  <p className="rounded-md bg-muted p-3">현재 상태 · {selected.transaction.status}</p>
                  {selected.transaction.answer && <p className="rounded-md bg-muted p-3">답변 · {selected.transaction.answer}</p>}
                </div>
              </div>
            </>
          ) : (
            <div className="rounded-lg border bg-card p-8 text-center text-sm text-muted-foreground">
              검토할 오류가 없습니다.
            </div>
          )}
        </aside>
      </section>
    </main>
  );
}
