# 💰 FinGenius AI — Smart Personal Finance Tracker

<div align="center">

![FinGenius AI](https://img.shields.io/badge/FinGenius-AI-6366f1?style=for-the-badge&logo=data:image/svg+xml;base64,...)
![MERN Stack](https://img.shields.io/badge/Stack-MERN-22c55e?style=for-the-badge)
![AI Powered](https://img.shields.io/badge/AI-Gemini-f59e0b?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-blue?style=for-the-badge)

**A production-ready AI-powered personal finance tracker built with the MERN stack.**  
Track expenses, manage budgets, set goals, and get AI-powered financial insights.

</div>

---

## ✨ Features

| Feature | Description |
|---|---|
| 🔐 **Authentication** | JWT-based auth with login, signup, forgot/reset password |
| 📊 **Dashboard** | Real-time overview of balance, income, expenses, savings |
| 💸 **Expense Tracking** | Add/edit/delete expenses with categories, receipts & location |
| 💰 **Income Tracking** | Track salary, freelancing, investments with recurring support |
| 📋 **Budget Planning** | Category-wise monthly budgets with alerts |
| 🎯 **Goal Planning** | Savings goals with AI-estimated completion dates |
| 🧾 **Receipt OCR** | Scan receipts to auto-extract amount, merchant, date |
| 🎤 **Voice Entry** | Speak to log expenses hands-free |
| 🤖 **AI Chat** | Gemini-powered assistant with access to your financial data |
| 📈 **Analytics** | 6 chart types: trends, pie, bar, area, comparison |
| 🔔 **Notifications** | Budget alerts, goal achievements, reminders |
| 🌙 **Dark/Light Mode** | Beautiful glassmorphism UI with theme toggle |
| 📱 **Responsive** | Works perfectly on mobile, tablet, and desktop |

---

## 🛠 Tech Stack

**Frontend**
- React 18 + Vite
- Tailwind CSS (glassmorphism design)
- Framer Motion (animations)
- Recharts (charts)
- React Hook Form
- React Router v6
- Axios

**Backend**
- Node.js + Express.js
- MongoDB + Mongoose
- JWT Authentication
- bcryptjs
- Multer (file uploads)
- Tesseract.js (OCR)
- Google Gemini AI

---

## 🚀 Getting Started

### Prerequisites

- Node.js >= 18.0.0
- MongoDB Atlas account (free tier works great)
- Google Gemini API key ([get one free](https://aistudio.google.com/app/apikey))

---

### 1. Clone & Install

```bash
# Clone the repo
git clone https://github.com/your-username/fingenius-ai.git
cd fingenius-ai

# Install backend dependencies
cd backend
npm install

# Install frontend dependencies
cd ../client
npm install
```

---

### 2. Configure Environment Variables

**Backend** — Copy and fill in your values:

```bash
cd backend
cp .env.example .env
```

Edit `backend/.env`:

```env
PORT=5000
NODE_ENV=development
MONGODB_URI=mongodb+srv://<user>:<pass>@<cluster>.mongodb.net/fingenius
JWT_SECRET=your_32_char_random_secret_here
JWT_EXPIRE=7d
JWT_RESET_SECRET=another_32_char_random_secret
CLIENT_URL=http://localhost:5173
GEMINI_API_KEY=your_gemini_api_key_here
```

> **Getting a Gemini API Key:**  
> 1. Go to [Google AI Studio](https://aistudio.google.com/app/apikey)  
> 2. Click "Create API Key"  
> 3. Copy the key into `GEMINI_API_KEY`

> **Getting MongoDB Atlas URI:**  
> 1. Create a free cluster at [mongodb.com/atlas](https://www.mongodb.com/atlas)  
> 2. Create a database user  
> 3. Whitelist your IP (or use 0.0.0.0/0 for dev)  
> 4. Copy the connection string

---

### 3. Start the Development Servers

**Terminal 1 — Backend:**

```bash
cd backend
npm run dev
# Server starts on http://localhost:5000
```

**Terminal 2 — Frontend:**

```bash
cd client
npm run dev
# App opens on http://localhost:5173
```

Open **http://localhost:5173** in your browser. Register an account and you're good to go!

---

## 📁 Project Structure

```
fingenius-ai/
├── backend/
│   ├── config/
│   │   ├── db.js              # MongoDB connection
│   │   ├── jwt.js             # JWT helpers
│   │   └── multer.js          # File upload config
│   ├── controllers/
│   │   ├── authController.js
│   │   ├── expenseController.js
│   │   ├── incomeController.js
│   │   ├── budgetController.js
│   │   ├── goalController.js
│   │   ├── dashboardController.js
│   │   ├── analyticsController.js
│   │   ├── chatController.js
│   │   ├── ocrController.js
│   │   ├── profileController.js
│   │   └── notificationController.js
│   ├── middleware/
│   │   ├── auth.js            # JWT protect middleware
│   │   ├── errorHandler.js    # Global error handler
│   │   ├── validate.js        # express-validator rules
│   │   └── rateLimiter.js     # Rate limiting
│   ├── models/
│   │   ├── User.js
│   │   ├── Expense.js
│   │   ├── Income.js
│   │   ├── Budget.js
│   │   ├── Goal.js
│   │   ├── Chat.js
│   │   └── Notification.js
│   ├── routes/
│   │   ├── auth.js
│   │   ├── expenses.js
│   │   ├── income.js
│   │   ├── budget.js
│   │   ├── goals.js
│   │   ├── dashboard.js
│   │   ├── analytics.js
│   │   ├── chat.js
│   │   ├── ocr.js
│   │   ├── profile.js
│   │   └── notifications.js
│   ├── services/
│   │   ├── aiService.js       # Gemini AI integration
│   │   ├── ocrService.js      # Tesseract OCR
│   │   └── emailService.js    # Nodemailer
│   ├── uploads/               # Uploaded files (gitignored)
│   ├── server.js              # Express app entry
│   ├── package.json
│   └── .env.example
│
└── client/
    ├── src/
    │   ├── components/
    │   │   ├── common/        # Shared: Sidebar, TopBar, Layout, etc.
    │   │   ├── expenses/      # ExpenseModal
    │   │   ├── income/        # IncomeModal
    │   │   ├── budget/        # BudgetModal
    │   │   └── goals/         # GoalModal, ContributeModal
    │   ├── context/
    │   │   ├── AuthContext.jsx
    │   │   └── ThemeContext.jsx
    │   ├── hooks/
    │   │   ├── useFetch.js
    │   │   ├── useVoiceInput.js
    │   │   ├── useDebounce.js
    │   │   └── useLocalStorage.js
    │   ├── pages/
    │   │   ├── Landing.jsx
    │   │   ├── Login.jsx
    │   │   ├── Register.jsx
    │   │   ├── ForgotPassword.jsx
    │   │   ├── ResetPassword.jsx
    │   │   ├── Dashboard.jsx
    │   │   ├── Expenses.jsx
    │   │   ├── Income.jsx
    │   │   ├── Budget.jsx
    │   │   ├── Goals.jsx
    │   │   ├── Analytics.jsx
    │   │   ├── Chat.jsx
    │   │   ├── Settings.jsx
    │   │   └── NotFound.jsx
    │   ├── services/          # Axios API service modules
    │   ├── utils/             # formatters, constants, helpers
    │   ├── App.jsx
    │   ├── main.jsx
    │   └── index.css
    ├── index.html
    ├── vite.config.js
    ├── tailwind.config.js
    ├── postcss.config.js
    └── package.json
```

---

## 🔌 API Reference

| Method | Endpoint | Description | Auth |
|--------|----------|-------------|------|
| POST | `/api/auth/register` | Register new user | No |
| POST | `/api/auth/login` | Login | No |
| GET | `/api/auth/me` | Get current user | Yes |
| POST | `/api/auth/forgot-password` | Send reset email | No |
| POST | `/api/auth/reset-password/:token` | Reset password | No |
| GET | `/api/dashboard` | Full dashboard summary | Yes |
| GET/POST | `/api/expenses` | List/create expenses | Yes |
| PUT/DELETE | `/api/expenses/:id` | Update/delete expense | Yes |
| GET/POST | `/api/income` | List/create income | Yes |
| GET/POST | `/api/budget` | List/create budgets | Yes |
| GET | `/api/budget/current` | Current month budget | Yes |
| GET/POST | `/api/goals` | List/create goals | Yes |
| POST | `/api/goals/:id/contribute` | Add contribution | Yes |
| POST | `/api/chat` | AI chat message | Yes |
| POST | `/api/ocr/receipt` | Process receipt image | Yes |
| GET | `/api/analytics/*` | All analytics endpoints | Yes |
| GET/PUT | `/api/profile` | Get/update profile | Yes |
| GET | `/api/notifications` | Get notifications | Yes |

---

## 🤖 AI Features Setup

### Gemini AI Chat
The AI assistant uses Google Gemini 1.5 Flash. It automatically receives:
- Monthly expense summary by category
- Income overview
- Budget status
- Active goals progress
- Last 5 transactions

It can answer questions like:
- *"How much did I spend on food this month?"*
- *"Can I afford a ₹70,000 laptop?"*
- *"Suggest budget improvements"*
- *"Show my unnecessary expenses"*

### Receipt OCR
Uses Tesseract.js to extract:
- Merchant name
- Amount
- Date
- GST
- Auto-detect category

### Voice Entry
Uses Web Speech API (Chrome/Edge). Say:
- *"I spent 250 rupees on petrol"*
- *"Paid 1500 for groceries"*

---

## 🔒 Security Features

- JWT authentication with expiry
- Password hashing with bcrypt (12 salt rounds)
- Rate limiting (100 req/15min, 10 auth/15min)
- CORS restricted to client URL
- Helmet.js security headers
- Input validation with express-validator
- MongoDB injection prevention via Mongoose

---

## 🚢 Production Deployment

### Backend (Render / Railway / Fly.io)

```bash
cd backend
# Set all .env variables in your hosting platform
npm start
```

### Frontend (Vercel / Netlify)

```bash
cd client
npm run build
# Deploy the dist/ folder
```

> **Important:** Update `CLIENT_URL` in backend `.env` and update the API proxy in `vite.config.js` with your production backend URL.

---

## 📦 Scripts

| Command | Description |
|---------|-------------|
| `cd backend && npm run dev` | Start backend in dev mode (nodemon) |
| `cd backend && npm start` | Start backend in production |
| `cd client && npm run dev` | Start frontend dev server |
| `cd client && npm run build` | Build frontend for production |
| `cd client && npm run preview` | Preview production build |

---

## 🤝 Contributing

1. Fork the repo
2. Create your feature branch: `git checkout -b feature/amazing-feature`
3. Commit changes: `git commit -m 'Add amazing feature'`
4. Push: `git push origin feature/amazing-feature`
5. Open a Pull Request

---

## 📄 License

MIT License — feel free to use this project for personal or commercial purposes.

---

<div align="center">

Built with ❤️ for India 🇮🇳  
**FinGenius AI** — *Your money, managed by AI*

</div>
