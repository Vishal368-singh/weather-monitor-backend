module.exports = {
  apps: [
    {
      name: "indus-weather-api",
      script: "waitress-serve",
      args: "--host=0.0.0.0 --port=6633 app:app",
      interpreter: "C:/inetpub/PM2-APIs/weather_FlaskAPI/py-env/Scripts/pythonw.exe"
    }
  ]
}
