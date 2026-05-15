#!/usr/bin/env bash
# Comercial JL Abogados (jlabogados.cl) — 30s en 6 clips de 5s
# Genera cada clip vía Higgsfield CLI y guarda las URLs en clip-urls.txt
#
# Uso:
#   1) npm install -g @higgsfield/cli
#   2) higgsfield auth login   (abre el navegador en tu máquina local)
#   3) ./generate-clips.sh [--model veo3_fast] [--aspect 16:9]
#
# Modelos sugeridos (lista completa: higgsfield model list --video):
#   veo3_fast    — Google Veo 3 rápido (recomendado para iterar)
#   veo3         — Google Veo 3 calidad
#   sora2_pro    — OpenAI Sora 2 Pro
#   kling_v2     — Kling v2
#   seedance_2   — Seedance 2.0

set -euo pipefail

MODEL="${MODEL:-veo3_fast}"
ASPECT="${ASPECT:-16:9}"
DURATION="${DURATION:-5}"
OUT_FILE="clip-urls.txt"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --model)  MODEL="$2";  shift 2 ;;
    --aspect) ASPECT="$2"; shift 2 ;;
    --duration) DURATION="$2"; shift 2 ;;
    *) echo "Flag desconocida: $1" >&2; exit 1 ;;
  esac
done

command -v higgsfield >/dev/null || { echo "higgsfield CLI no instalado. npm install -g @higgsfield/cli" >&2; exit 1; }
higgsfield auth token >/dev/null 2>&1 || { echo "No autenticado. Corre: higgsfield auth login" >&2; exit 1; }

declare -a PROMPTS=(
  # CLIP 1 — Apertura cinematográfica (dolly push-in)
  "Black and white cinematic 35mm portrait of a confident lawyer wearing a tailored dark suit and white construction hardhat with sunglasses, leaning on metal scaffolding railing on an active construction site, vertical rebar rods rising in the background, soft overcast natural light, slow dolly push-in toward his face, shallow depth of field, anamorphic lens flare, subtle film grain, premium luxury law firm commercial, monochrome with deep blacks"

  # CLIP 2 — Mirada y autoridad (slow orbit)
  "Slow cinematic orbit around the same confident lawyer in tailored dark suit on a construction site, he calmly removes his sunglasses and looks directly at camera with composed confidence, dust particles drifting in the air, dramatic side light, monochrome with warm gold highlights, anamorphic 35mm, high-end law firm commercial"

  # CLIP 3 — Firma de contrato (crash zoom)
  "Extreme close-up of a luxury fountain pen signing a construction contract on a dark wooden desk, architectural blueprints and a small modern building model beside it, warm tungsten light, fast crash zoom into the signature line, cinematic, sophisticated, premium 35mm film grain, shallow depth of field"

  # CLIP 4 — Áreas de servicio (crane up sobre edificio)
  "Cinematic crane shot rising up alongside a modern high-rise building under construction at golden hour, workers and tall cranes in the distance, monochrome with subtle gold tint, epic, prestigious corporate aesthetic, anamorphic 35mm, slow upward camera motion"

  # CLIP 5 — Mensaje central (typography reveal)
  "Cinematic deep black background with elegant gold serif typography slowly revealing the phrase 'Estrategia legal a la altura de tu proyecto', soft gold particles drifting upward, premium law firm brand identity, minimalist luxury, subtle parallax camera motion, anamorphic flare"

  # CLIP 6 — Cierre con CTA (logo reveal)
  "Elegant centered logo reveal on deep black background, sophisticated thin serif wordmark 'JL ABOGADOS' fading in with a thin gold underline animating beneath, then the line 'Agenda tu asesoría en www.jlabogados.cl' appears in smaller type, premium fade out, cinematic, minimalist luxury"
)

declare -a TITLES=(
  "01-apertura-dolly-in"
  "02-orbit-mirada"
  "03-firma-crash-zoom"
  "04-crane-up-edificio"
  "05-typography-mensaje"
  "06-cta-logo-reveal"
)

: > "$OUT_FILE"
echo "Modelo:   $MODEL"
echo "Aspecto:  $ASPECT"
echo "Duración: ${DURATION}s por clip"
echo "Total clips: ${#PROMPTS[@]}"
echo

for i in "${!PROMPTS[@]}"; do
  n=$((i + 1))
  title="${TITLES[$i]}"
  prompt="${PROMPTS[$i]}"
  echo "===================================================="
  echo "[$n/6] $title"
  echo "===================================================="

  if ! higgsfield generate create "$MODEL" \
      --prompt "$prompt" \
      --aspect_ratio "$ASPECT" \
      --duration "$DURATION" \
      --wait --wait-timeout 15m --wait-interval 10s \
      --json > "clip-${n}.json" 2>&1; then
    echo "FALLO en clip $n — ver clip-${n}.json" >&2
    cat "clip-${n}.json" >&2
    continue
  fi

  url=$(jq -r '.. | objects | select(has("url")) | .url' "clip-${n}.json" 2>/dev/null | head -1 || true)
  [[ -z "$url" || "$url" == "null" ]] && url=$(grep -oE 'https://[^"[:space:]]+\.mp4[^"[:space:]]*' "clip-${n}.json" | head -1 || true)

  if [[ -n "$url" ]]; then
    echo "OK: $url"
    printf "%s\t%s\t%s\n" "$n" "$title" "$url" >> "$OUT_FILE"
  else
    echo "Clip $n terminó pero no se detectó URL — revisa clip-${n}.json"
  fi
done

echo
echo "===================================================="
echo "URLs guardadas en $OUT_FILE"
column -t -s $'\t' "$OUT_FILE" 2>/dev/null || cat "$OUT_FILE"
echo
echo "Montaje sugerido:"
echo "  Descarga los 6 MP4 y únelos en CapCut/Premiere/DaVinci en orden 1->6."
echo "  Añade música cinematic corporate (BPM 80-90), voz en off opcional, LUT B/N."
