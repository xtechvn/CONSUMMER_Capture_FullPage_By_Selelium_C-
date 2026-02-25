# AGENTS.md

## Cursor Cloud specific instructions

### Overview

ConsummerScreenPageBot is a .NET 8.0 console application that consumes jobs from RabbitMQ, uses Selenium + ChromeDriver to navigate to Vietnamese news websites, takes full-page screenshots, segments them, and publishes results to RabbitMQ analysis queues. It also detects ad banners/iframes and captures landing pages.

### Prerequisites (installed by VM snapshot)

- **.NET 8.0 SDK** — installed at `/usr/share/dotnet`
- **Google Chrome** — pre-installed on the VM
- **Docker** — required to run RabbitMQ locally (fuse-overlayfs + iptables-legacy configured)

### Running RabbitMQ locally

The app requires RabbitMQ. Start Docker and RabbitMQ before running:

```bash
sudo dockerd &>/tmp/dockerd.log &
sleep 3
sudo docker start rabbitmq 2>/dev/null || \
  sudo docker run -d --name rabbitmq -p 5672:5672 -p 15672:15672 \
    -e RABBITMQ_DEFAULT_USER=guest -e RABBITMQ_DEFAULT_PASS=guest \
    rabbitmq:3-management
sudo docker exec rabbitmq rabbitmqctl await_startup
```

### App.config for local development

Before running, update `ConsummerScreenPageBot/App.config` to point to local RabbitMQ:

| Key | Production value | Local dev value |
|-----|-----------------|-----------------|
| `RabbitHost` | `103.163.216.115` | `localhost` |
| `RabbitUserName` | `web_push` | `guest` |
| `RabbitPassword` | `123465` | `guest` |
| `RabbitVHost` | `booking_car` | `/` |
| `is_headless` | `0` | `1` (no display server in cloud VM) |

### Build and run

```bash
cd ConsummerScreenPageBot
dotnet restore
dotnet build
dotnet run
```

### Sending a test message

Use `rabbitmqadmin` inside the container to publish a test job:

```bash
sudo docker exec rabbitmq rabbitmqadmin publish \
  exchange=amq.default \
  routing_key=QUEUE_PROCESS_IMAGE_SCREEN_TEST \
  payload='{"link_web":"https://vnexpress.net","slice":3,"quanlity_image":70,"device":"1","retry_screen_page":1}'
```

### Lint

```bash
dotnet format ConsummerScreenPageBot.csproj --verify-no-changes
```

Note: The codebase has pre-existing whitespace formatting issues. Use the `.csproj` path explicitly since both `.sln` and `.csproj` exist in the same directory.

### Key gotchas

- **No test framework**: There are no unit/integration tests in this codebase.
- **ChromeDriver version**: The NuGet package `Selenium.WebDriver.ChromeDriver` bundles chromedriver, but Chrome on the VM auto-updates. If Chrome and chromedriver versions diverge significantly, Selenium will fail to start. The current setup works with Chrome 145.x.
- **Screenshots output**: Screenshots are saved to `bin/Debug/net8.0/screenshots/<hostname>/<device>/`.
- **App runs indefinitely**: The bot enters an infinite wait loop listening for RabbitMQ messages. Use `timeout` or Ctrl-C to stop it.
- **Revert App.config before committing**: Always revert local RabbitMQ config changes before pushing.
