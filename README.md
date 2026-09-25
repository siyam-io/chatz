# 💬 Chat-Z

Welcome to **Chat-Z**, a comprehensive, real-time messaging ecosystem. This project features a robust mobile application, a scalable PostgreSQL backend, and a dedicated Next.js administrative dashboard. 

![Chat-Z Banner](https://via.placeholder.com/1200x300.png?text=Chat-Z+-+Real-Time+Messaging+Ecosystem)

## 🌟 Features

- **Real-Time Messaging**: Lightning-fast socket-based communication.
- **Mobile First**: Built with React Native & Expo for cross-platform (iOS/Android) mobile capabilities.
- **Secure Backend**: Express & Node.js API with PostgreSQL and structured schemas.
- **Admin Dashboard**: A sleek, feature-rich Next.js dashboard to manage users, data, and configurations.
- **Scalable Architecture**: Well-separated components ensuring maintainability and ease of scaling.

---

## 🏗 Project Architecture

The project is divided into three main components:

### 1. 📱 ChatApp (Mobile App)
- **Tech Stack**: React Native, Expo
- **Description**: The core user-facing mobile application where users can communicate seamlessly.

### 2. ⚙️ Backend (`backend-pg`)
- **Tech Stack**: Node.js, Express, PostgreSQL, Socket.io
- **Description**: The engine of Chat-Z. It handles real-time socket connections, RESTful APIs, data persistence, and core business logic.

### 3. 🛡 Admin (`admin`)
- **Tech Stack**: Next.js, React
- **Description**: The administrative interface for managing the platform, users, and monitoring activities.

---

## 🚀 Getting Started

### Prerequisites

- Node.js (v18 or higher)
- PostgreSQL
- Expo CLI

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/your-username/chat-z.git
   cd chat-z
   ```

2. **Backend Setup:**
   ```bash
   cd backend-pg
   npm install
   # Copy .env.example to .env and configure your PostgreSQL database
   cp .env.example .env
   # Run migrations
   npm run db:migrate
   # Start the server
   npm run dev
   ```

3. **Mobile App Setup:**
   ```bash
   cd ../ChatApp
   npm install
   # Start the Expo server
   npx expo start
   ```

4. **Admin Dashboard Setup:**
   ```bash
   cd ../admin
   npm install
   # Start the development server
   npm run dev
   ```

---

## 🔒 Environment Variables

Make sure to create `.env` files in their respective directories (`backend-pg`, `ChatApp`, `admin`). Refer to `.env.example` in each directory for required variables.

---

## 🛠 Built With

* [React Native](https://reactnative.dev/) & [Expo](https://expo.dev/)
* [Next.js](https://nextjs.org/)
* [Node.js](https://nodejs.org/) & [Express](https://expressjs.com/)
* [PostgreSQL](https://www.postgresql.org/)
* [Socket.io](https://socket.io/)

---

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.
