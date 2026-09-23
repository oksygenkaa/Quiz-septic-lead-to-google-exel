import { createFileRoute } from "@tanstack/react-router";
import { useCallback, useEffect, useRef, useState } from "react";

/* =============================================================================
 * Квиз-лендинг «Времена года» — имитация переписки в мессенджере.
 *
 * ---------------------------------------------------------------------------
 * ШАБЛОН Google Apps Script (Code.gs)
 * Вставьте в таблицу: Расширения → Apps Script, затем Deploy → New deployment
 * → Web app → Execute as: Me, Who has access: Anyone. Полученный URL
 * подставьте в константу GOOGLE_SCRIPT_URL ниже.
 *
 * function doPost(e) {
 *   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
 *   var p = e.parameter || {};
 *   if (sheet.getLastRow() === 0) {
 *     sheet.appendRow([
 *       'timestamp','answer_1','answer_2','answer_3','answer_4','name','phone',
 *       'utm_source','utm_medium','utm_campaign','utm_content','utm_term','page_url'
 *     ]);
 *   }
 *   sheet.appendRow([
 *     p.timestamp || new Date(),
 *     p.answer_1 || '', p.answer_2 || '', p.answer_3 || '', p.answer_4 || '',
 *     p.name || '', p.phone || '',
 *     p.utm_source || '', p.utm_medium || '', p.utm_campaign || '',
 *     p.utm_content || '', p.utm_term || '', p.page_url || ''
 *   ]);
 *   return ContentService.createTextOutput('OK');
 * }
 * ---------------------------------------------------------------------------
 */

const GOOGLE_SCRIPT_URL = "ВСТАВЬТЕ_СЮДА_URL_ВАШЕГО_APPS_SCRIPT";
const CALL_MARKING = "IP Nikitina IV: stroit-vo";

const UTM_KEYS = [
  "utm_source",
  "utm_medium",
  "utm_campaign",
  "utm_content",
  "utm_term",
] as const;

type UtmKey = (typeof UTM_KEYS)[number];
type Utm = Record<UtmKey, string>;

const EMPTY_UTM: Utm = {
  utm_source: "",
  utm_medium: "",
  utm_campaign: "",
  utm_content: "",
  utm_term: "",
};

type Step = {
  messages: string[];
  options: string[];
};

const STEPS: Step[] = [
  {
    messages: [
      "Здравствуйте! Я Максим, менеджер компании «Времена года».",
      "Подберём оптимальный септик под ваш дом и бюджет — а не тот, что выгоднее продать. Цены от производителя — от 55 000 рублей. Гарантия до 50 лет и оплата по факту.",
      "Ответьте на 3 коротких вопроса — подготовлю расчёт для вас.",
      "1. Сколько человек проживает или будет проживать в доме?",
    ],
    options: ["От 1 до 3 человек", "4–6 человек", "7–11 человек", "Больше 12 человек"],
  },
  {
    messages: ["2. Как планируете проживать на объекте?"],
    options: ["Сезонно", "Постоянно", "Коммерческий объект"],
  },
  {
    messages: ["3. Укажите уровень грунтовых вод на Вашем участке?"],
    options: ["Низкий, Ниже 1.5 метра", "Высокий, Выше 1.5 метра", "Не знаю"],
  },
  {
    messages: ["4. Как Вам удобнее получить подборку септика?"],
    options: ["По телефону", "Telegram", "MAX", "WhatsApp"],
  },
  {
    messages: [
      "Спасибо! Расчёт практически готов. Оставьте номер телефона, мы уточним несколько вопросов и рассчитаем стоимость под ключ.",
      "Осень — подходящее время для установки септика: основной сезон заканчивается, участок меньше используется, а монтаж можно провести до промерзания грунта и подготовить канализацию к следующему сезону.",
      "Укажите Ваше имя и номер телефона для связи выбранным способом.",
    ],
    options: [],
  },
];

const TOTAL_STEPS = STEPS.length;

type ChatItem = { id: number; from: "manager" | "user"; text: string };

function formatPhone(raw: string): string {
  let digits = raw.replace(/\D/g, "");
  if (digits.startsWith("8")) digits = "7" + digits.slice(1);
  if (!digits.startsWith("7")) digits = "7" + digits;
  digits = digits.slice(0, 11);
  const rest = digits.slice(1);
  let out = "+7";
  if (rest.length > 0) out += ` (${rest.slice(0, 3)}`;
  if (rest.length >= 3) out += ")";
  if (rest.length > 3) out += ` ${rest.slice(3, 6)}`;
  if (rest.length > 6) out += `-${rest.slice(6, 8)}`;
  if (rest.length > 8) out += `-${rest.slice(8, 10)}`;
  return out;
}

function phoneComplete(value: string): boolean {
  return value.replace(/\D/g, "").length === 11;
}

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Времена года — подбор септика за 4 вопроса" },
      {
        name: "description",
        content:
          "Ответьте на 4 коротких вопроса — подберём септик под ваш дом и бюджет. Цены от производителя от 55 000 ₽, гарантия до 50 лет, оплата по факту.",
      },
      { property: "og:title", content: "Времена года — подбор септика за 4 вопроса" },
      {
        property: "og:description",
        content:
          "Подбор септика под ваш дом и бюджет: цены от производителя от 55 000 ₽, гарантия до 50 лет, оплата по факту.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: QuizPage,
});

function ManagerAvatar({ size = "sm" }: { size?: "sm" | "lg" }) {
  return (
    <div
      className={`grid shrink-0 place-items-center rounded-full bg-primary font-bold text-primary-foreground ${
        size === "lg" ? "h-11 w-11 text-lg" : "h-7 w-7 text-xs"
      }`}
      aria-hidden="true"
    >
      M
    </div>
  );
}

function TypingBubble() {
  return (
    <div className="flex items-end gap-2">
      <ManagerAvatar />
      <div className="flex items-center gap-1 rounded-2xl rounded-bl-md bg-chat-bubble px-4 py-3 shadow-sm">
        {[0, 1, 2].map((i) => (
          <span
            key={i}
            className="animate-dot inline-block h-1.5 w-1.5 rounded-full bg-muted-foreground"
            style={{ animationDelay: `${i * 0.18}s` }}
          />
        ))}
      </div>
    </div>
  );
}

function QuizPage() {
  const [stepIndex, setStepIndex] = useState(0);
  const [chat, setChat] = useState<ChatItem[]>([]);
  const [typing, setTyping] = useState(true);
  const [showInteraction, setShowInteraction] = useState(false);
  const [answers, setAnswers] = useState<string[]>([]);
  const [name, setName] = useState("");
  const [phone, setPhone] = useState("");
  const [agreed, setAgreed] = useState(false);
  const [done, setDone] = useState(false);
  const [modalOpen, setModalOpen] = useState(false);
  const [submittedPhone, setSubmittedPhone] = useState("");
  const [utm, setUtm] = useState<Utm>(EMPTY_UTM);

  const idRef = useRef(0);
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const stored = sessionStorage.getItem("quiz_utm");
    const params = new URLSearchParams(window.location.search);
    const fromUrl = { ...EMPTY_UTM };
    let hasAny = false;
    for (const key of UTM_KEYS) {
      const value = params.get(key);
      if (value) {
        fromUrl[key] = value;
        hasAny = true;
      }
    }
    if (hasAny) {
      sessionStorage.setItem("quiz_utm", JSON.stringify(fromUrl));
      setUtm(fromUrl);
    } else if (stored) {
      try {
        setUtm({ ...EMPTY_UTM, ...(JSON.parse(stored) as Partial<Utm>) });
      } catch {
        setUtm(EMPTY_UTM);
      }
    }
  }, []);

  // Последовательный «набор» сообщений менеджера для текущего шага.
  useEffect(() => {
    const step = STEPS[stepIndex];
    if (!step) return;
    let cancelled = false;
    const timers: ReturnType<typeof setTimeout>[] = [];

    const play = (i: number) => {
      if (cancelled) return;
      const text = step.messages[i];
      if (text === undefined) {
        timers.push(
          setTimeout(() => {
            if (!cancelled) setShowInteraction(true);
          }, 350),
        );
        return;
      }
      setTyping(true);
      const delay = Math.min(2000, 900 + text.length * 8);
      timers.push(
        setTimeout(() => {
          if (cancelled) return;
          idRef.current += 1;
          setChat((prev) => [...prev, { id: idRef.current, from: "manager", text }]);
          setTyping(false);
          timers.push(
            setTimeout(() => {
              play(i + 1);
            }, 400),
          );
        }, delay),
      );
    };

    play(0);
    return () => {
      cancelled = true;
      timers.forEach(clearTimeout);
    };
  }, [stepIndex]);

  useEffect(() => {
    const el = scrollRef.current;
    if (el) el.scrollTo({ top: el.scrollHeight, behavior: "smooth" });
  }, [chat, typing, showInteraction, done]);

  const handleOption = (option: string) => {
    idRef.current += 1;
    setChat((prev) => [...prev, { id: idRef.current, from: "user", text: option }]);
    setAnswers((prev) => [...prev, option]);
    setShowInteraction(false);
    setTyping(true);
    setStepIndex((prev) => prev + 1);
  };

  const canSubmit = name.trim().length > 0 && phoneComplete(phone) && agreed;

  const handleSubmit = useCallback(async () => {
    if (!canSubmit) return;
    const payload = {
      timestamp: new Date().toISOString(),
      answer_1: answers[0] ?? "",
      answer_2: answers[1] ?? "",
      answer_3: answers[2] ?? "",
      answer_4: answers[3] ?? "",
      name: name.trim(),
      phone,
      ...utm,
      page_url: window.location.href,
    };

    if (GOOGLE_SCRIPT_URL.startsWith("http")) {
      try {
        await fetch(GOOGLE_SCRIPT_URL, {
          method: "POST",
          mode: "no-cors",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams(payload).toString(),
        });
      } catch (error) {
        console.error("Ошибка отправки заявки", error);
      }
    } else {
      console.log("Данные заявки (URL Apps Script не задан):", payload);
    }

    setSubmittedPhone(phone);
    setShowInteraction(false);
    setDone(true);
    setTimeout(() => setModalOpen(true), 800);
  }, [answers, canSubmit, name, phone, utm]);

  const stepNumber = Math.min(stepIndex + 1, TOTAL_STEPS);
  const progress = (stepNumber / TOTAL_STEPS) * 100;
  const currentStep = STEPS[stepIndex];
  const isFormStep = stepIndex === TOTAL_STEPS - 1;

  return (
    <main className="flex min-h-screen justify-center bg-background md:items-center md:p-6">
      <div className="flex h-[100dvh] w-full flex-col overflow-hidden bg-card shadow-xl md:h-[820px] md:max-h-[92vh] md:w-[440px] md:rounded-3xl md:border md:border-border">
        {/* Шапка */}
        <header className="shrink-0 bg-card pt-[env(safe-area-inset-top)]">
          <div className="grid grid-cols-[auto_minmax(0,1fr)] items-center gap-3 px-4 py-3">
            <ManagerAvatar size="lg" />
            <div className="min-w-0">
              <p className="truncate font-semibold text-card-foreground">Менеджер Максим</p>
              <p className="flex items-center gap-1.5 text-xs text-muted-foreground">
                {!done && <span className="h-1.5 w-1.5 rounded-full bg-success" />}
                {done ? "Готово!" : "Онлайн"}
              </p>
            </div>
          </div>
          <div className="border-t border-border px-4 pb-3 pt-2">
            <p className="text-xs text-muted-foreground">
              Шаг {stepNumber} из {TOTAL_STEPS}
            </p>
            <div className="mt-1.5 h-[3px] w-full overflow-hidden rounded-full bg-muted">
              <div
                className="h-full rounded-full bg-quiz-cta transition-all duration-700 ease-out"
                style={{ width: `${Math.min(100, progress)}%` }}
              />
            </div>
          </div>
        </header>

        {/* Тело чата */}
        <div ref={scrollRef} className="flex-1 space-y-3 overflow-y-auto bg-chat-surface px-4 py-4">
          {chat.map((item) =>
            item.from === "manager" ? (
              <div key={item.id} className="animate-bubble-in flex items-end gap-2">
                <ManagerAvatar />
                <div className="max-w-[80%] rounded-2xl rounded-bl-md bg-chat-bubble px-4 py-2.5 text-[15px] leading-snug text-chat-bubble-foreground shadow-sm">
                  {item.text}
                </div>
              </div>
            ) : (
              <div key={item.id} className="animate-bubble-in flex justify-end">
                <div className="max-w-[80%] rounded-2xl rounded-br-md bg-chat-user-bubble px-4 py-2.5 text-[15px] leading-snug text-chat-bubble-foreground shadow-sm">
                  {item.text}
                </div>
              </div>
            ),
          )}

          {typing && !done && <TypingBubble />}

          {/* Финальный экран */}
          {done && (
            <div className="animate-bubble-in flex flex-col items-center gap-3 px-2 py-8 text-center">
              <div className="animate-pop-in grid h-16 w-16 place-items-center rounded-full bg-success text-success-foreground">
                <svg
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth={3}
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  className="h-8 w-8"
                  aria-hidden="true"
                >
                  <path d="M20 6 9 17l-5-5" />
                </svg>
              </div>
              <h2 className="text-xl font-bold text-foreground">Спасибо, заявка принята!</h2>
              <p className="text-sm leading-relaxed text-muted-foreground">
                Звонок поступит с маркировкой «{CALL_MARKING}». Наш менеджер свяжется с вами в
                ближайшее время по этому номеру телефона
              </p>
              <p className="text-2xl font-bold text-primary">{submittedPhone}</p>
            </div>
          )}
        </div>

        {/* Варианты ответа */}
        {showInteraction && !done && !isFormStep && currentStep && (
          <div className="animate-bubble-in shrink-0 space-y-2 bg-card px-4 py-3">
            {currentStep.options.map((option) => (
              <button
                key={option}
                type="button"
                onClick={() => handleOption(option)}
                className="w-full min-h-11 rounded-xl border border-input bg-card px-4 py-3 text-left text-[15px] font-medium text-card-foreground shadow-sm transition-all hover:border-primary hover:bg-accent active:scale-[0.99]"
              >
                {option}
              </button>
            ))}
          </div>
        )}

        {/* Форма контактов */}
        {showInteraction && !done && isFormStep && (
          <div className="animate-bubble-in shrink-0 space-y-3 bg-card px-4 py-3">
            <input
              type="text"
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="Имя"
              aria-label="Ваше имя"
              maxLength={60}
              className="w-full min-h-11 rounded-xl border border-input bg-card px-4 py-2.5 text-[15px] outline-none transition-colors focus:border-primary"
            />
            <div className="flex min-h-11 items-center gap-2 rounded-xl border border-input bg-card px-3 transition-colors focus-within:border-primary">
              <span className="shrink-0 text-lg" aria-hidden="true">
                🇷🇺
              </span>
              <input
                type="tel"
                value={phone}
                onChange={(e) => setPhone(formatPhone(e.target.value))}
                onFocus={() => {
                  if (!phone) setPhone("+7");
                }}
                placeholder="+7 (999) 000-00-00"
                aria-label="Номер телефона"
                className="min-w-0 flex-1 bg-transparent py-2.5 text-[15px] outline-none"
              />
            </div>
            <label className="flex cursor-pointer gap-2.5 text-xs leading-relaxed text-muted-foreground">
              <input
                type="checkbox"
                checked={agreed}
                onChange={(e) => setAgreed(e.target.checked)}
                className="mt-0.5 h-4 w-4 shrink-0 accent-primary"
              />
              <span>
                Я согласен с Политикой конфиденциальности и Согласием на обработку персональных
                данных
              </span>
            </label>
            <button
              type="button"
              onClick={handleSubmit}
              disabled={!canSubmit}
              className="w-full min-h-12 rounded-xl bg-quiz-cta px-4 text-[15px] font-semibold text-foreground transition-all hover:opacity-90 active:scale-[0.99] disabled:cursor-not-allowed disabled:opacity-40"
            >
              Получить расчёт со скидкой →
            </button>
          </div>
        )}

        {/* Футер-дисклеймер */}
        <footer className="shrink-0 overflow-hidden border-t border-border bg-card px-2 py-3 pb-[max(0.75rem,env(safe-area-inset-bottom))] text-center text-[10px] leading-relaxed text-muted-foreground md:px-4 md:text-[11px]">
          <span className="inline-flex flex-col items-center gap-0.5 text-[10px] md:flex-row md:flex-nowrap md:gap-1.5 md:whitespace-nowrap md:tracking-normal">
            <span>Политика конфиденциальности</span>
            <span aria-hidden="true" className="hidden text-border md:inline">
              |
            </span>
            <span>Согласие на обработку персональных данных</span>
          </span>
        </footer>
      </div>

      {/* Модальное окно */}
      {modalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-foreground/50 px-4">
          <div className="animate-bubble-in w-full max-w-[340px] rounded-2xl bg-card p-6 text-center shadow-xl">
            <div className="animate-pop-in mx-auto grid h-14 w-14 place-items-center rounded-full bg-success text-success-foreground">
              <svg
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth={3}
                strokeLinecap="round"
                strokeLinejoin="round"
                className="h-7 w-7"
                aria-hidden="true"
              >
                <path d="M20 6 9 17l-5-5" />
              </svg>
            </div>
            <h3 className="mt-4 text-lg font-bold text-card-foreground">Спасибо!</h3>
            <p className="mt-1 text-sm text-muted-foreground">Данные успешно отправлены.</p>
            <button
              type="button"
              onClick={() => setModalOpen(false)}
              className="mt-5 w-full min-h-11 rounded-xl bg-foreground text-[15px] font-semibold text-background transition-opacity hover:opacity-90"
            >
              Хорошо
            </button>
          </div>
        </div>
      )}
    </main>
  );
}
