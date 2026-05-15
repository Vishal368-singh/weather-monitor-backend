# Weather Monitor Backend API

Backend API service for a real-time weather monitoring platform built using Flask.

## Features

- JWT Authentication
- Session-Based Authentication
- RESTful API Architecture
- Secure Route Protection
- Weather Data APIs
- User Authentication & Authorization
- Redis Caching Support

## Tech Stack

- Python
- Flask
- Flask-JWT-Extended
- Flask-Session
- MongoDB / SQL
- Redis

## Architecture

- Flask handles API routing and middleware
- JWT used for stateless authentication
- Session management implemented for secure login handling
- Redis used for caching and optimized response handling

## API Modules

- Authentication APIs
- Weather Data APIs
- User Management APIs
- Dashboard APIs

## Run Locally

```bash
pip install -r requirements.txt
python app.py
