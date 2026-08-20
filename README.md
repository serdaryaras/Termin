# ARTI İş Sıralama ve Gantt

Small Code Hub ile aynı görünüm ve yığın: **Next.js · Supabase · Vercel · GitHub**.

Excel’den iş listesi yapıştırılır, sürükle-bırak ile öncelik ve paralel yollar kurulur, Gantt otomatik üretilir.

## Yerelde çalıştırma

```bash
npm install
cp .env.example .env.local
npm run dev
```

[http://localhost:3000](http://localhost:3000)

## Supabase

1. [supabase.com](https://supabase.com) üzerinde bir proje açın.
2. SQL Editor’de `supabase/schema.sql` dosyasını çalıştırın.
3. Project Settings → API içinden URL ve anon key alın, `.env.local` (ve Vercel Environment Variables) içine yazın:

```
NEXT_PUBLIC_SUPABASE_URL=https://xxxx.supabase.co
NEXT_PUBLIC_SUPABASE_ANON_KEY=...
NEXT_PUBLIC_PLAN_SLUG=default
```

Anahtarlar olmadan uygulama açılır; kayıt yalnızca oturumda kalır.

## GitHub + Vercel

Small calculators deposundaki gibi:

```bash
git init
git add .
git commit -m "Initial Next.js Gantt app"
git remote add origin https://github.com/serdaryaras/arti-gantt-chart.git
git push -u origin main
```

Vercel → **Add New Project** → bu GitHub deposunu içe aktarın. Framework: **Next.js**. Aynı `NEXT_PUBLIC_*` değişkenlerini Vercel’e ekleyin. `main` dalına her push otomatik yayınlar.

## Klasörler

```
src/app            Sayfalar
src/components     Gantt arayüzü, ARTI logosu
src/lib            Sıralama, Excel ayrıştırma, Supabase
supabase/schema.sql
```
