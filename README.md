# Chat-Z

Chat-Z is a comprehensive, real-time messaging ecosystem designed for seamless communication. The project consists of a robust React Native mobile application, a scalable Node.js & PostgreSQL backend, and a feature-rich Next.js administrative dashboard.

## 🔗 Live Links

- **Admin Panel:** [https://chatz-iota-mocha.vercel.app/](https://chatz-iota-mocha.vercel.app/)
- **Backend API (Production):** [https://chaz-backend.onrender.com](https://chaz-backend.onrender.com)
- **Repository:** [https://github.com/siyam-io/chatz](https://github.com/siyam-io/chatz)

---

## 🌟 Key Features

- **Real-Time Messaging**: Lightning-fast socket-based communication ensuring low latency.
- **Cross-Platform Mobile App**: Built with React Native and Expo for both iOS and Android.
- **Robust Backend**: Node.js API with PostgreSQL for structured and secure data persistence.
- **Administrative Dashboard**: A modern Next.js interface for user management, analytics, and platform configuration.
- **Scalable Architecture**: Decoupled monolithic components for maintainability and horizontal scalability.

---

## 🏗 Project Architecture

The repository is structured into three primary components:

### 1. Mobile Application (`ChatApp`)
- **Framework**: React Native, Expo
- **Purpose**: The core user-facing client application designed for real-time chat interactions.

### 2. Backend Engine (`backend-pg`)
- **Framework**: Node.js, Express, PostgreSQL, Socket.io
- **Purpose**: Handles authentication, RESTful routing, real-time socket events, and database migrations.

### 3. Administrative Dashboard (`admin`)
- **Framework**: Next.js (React), Tailwind CSS
- **Purpose**: A secure web interface for platform administrators to monitor and manage the ecosystem.

---

## 🚀 Getting Started

### Prerequisites

Ensure you have the following installed on your local development environment:
- **Node.js** (v18.x or higher recommended)
- **PostgreSQL** (v14 or higher)
- **Expo CLI** (for mobile app development)

### Local Development Setup

#### 1. Repository Setup
```bash
git clone https://github.com/siyam-io/chatz.git
cd chatz
```

#### 2. Backend Initialization
```bash
cd backend-pg
npm install

# Environment configuration
cp .env.example .env
# Update the .env file with your local PostgreSQL credentials

# Execute database migrations
npm run db:migrate

# Start the development server (runs on port 5001 or 5002)
npm run dev
```

#### 3. Mobile Application Initialization
```bash
cd ../ChatApp
npm install

# Start the Expo development server
npx expo start
```

#### 4. Admin Dashboard Initialization
```bash
cd ../admin
npm install

# Configure environment variables
cp .env.example .env.local
# Update NEXT_PUBLIC_API_URL if needed

# Start the Next.js development server
npm run dev
```

---

## 🛠 Technology Stack

- **Frontend:** React Native, Expo, Next.js, React
- **Backend:** Node.js, Express, Socket.io
- **Database:** PostgreSQL
- **Deployment:** Vercel (Admin), Render (Backend)

---

## 📄 License

This project is proprietary. Please contact the repository owner for licensing details.
