# GYMZUP × Etsy — Digital Products

Kolchi dyal store Etsy kayn f had dossier:

| Fichier | Achno fih |
|---|---|
| `PLAN.md` | Niche, 10 produits, keywords, prix, strategy |
| `etsy_api.py` | Script kayrbet m3a Etsy API (listings, images, fichiers, commandes) |
| `products/<slug>/listing.json` | Titre, description, tags, prix dyal kol produit |
| `products/<slug>/*.pdf` + `images/` | Fichier li kaytbaa3 + tsawer dyal listing |

---

## 1. Kifach trbet Etsy m3a script (mra wa7da)

1. Sir l **https://www.etsy.com/developers/register** w sjjel app:
   - Name: `GYMZUP Tools` (wla li bghiti)
   - "Who will be the users?" → *Just myself or colleagues*
   - "Is your application commercial?" → *No*
2. F app dyalk, f **Callback URLs**, zid: `http://localhost:3003/callback`
3. Tsenna Etsy tvalidi l-app (status ywli *Active*). Ymken ykhod chi ayam.
4. Copy **Keystring** w **Shared secret**.
5. F terminal dyal PC dyalk (Python 3.9+, ma khass ta library):

```bash
export ETSY_KEYSTRING="xxxxxxxx"
export ETSY_SHARED_SECRET="yyyyyyyy"
python3 etsy/etsy_api.py auth
```

Navigateur ghadi y7el, dir **Allow access**, w tokens ytsjjlo f `etsy/.etsy_tokens.json`.
Had l-fichier secret, w kayn f `.gitignore`, ya3ni ma ytpushach l GitHub.

> ⚠️ Ma t7tt keystring wla secret wla tokens f chat wla f commit.

## 2. Commandes

```bash
python3 etsy/etsy_api.py me                       # yjib shop_id w smiya dyal store
python3 etsy/etsy_api.py taxonomy planner         # y9leb 3la taxonomy_id dyal catégorie
python3 etsy/etsy_api.py create etsy/products/12-week-workout-planner/listing.json
python3 etsy/etsy_api.py listings                 # listings dyalk (draft/active)
python3 etsy/etsy_api.py orders                   # akhir commandes
```

`create` kaycréé listing f **draft** (ma kaytpublishach direct). Kayuploadi tsawer w fichier digital,
w mn b3d nta tchouf kolchi f Etsy w tdir Publish b yeddik.

## 3. Achno khass t-tأكد mno

Docs rasmiya dyal Etsy ma kanetch accessible mn l-environment fin tktb had script, donc t-tktb
3la 7sab Open API v3 kif kan m3rouf. Ila tla3 lik chi error, sift lia l-message:
- Header `x-api-key`: script kaysifto `keystring:shared_secret` (Etsy wllat kat-tlb had format).
  Ila tla3 lik 401/403 f `me`, jrrb `export ETSY_KEY_FORMAT=keystring_only`.
- `taxonomy_id` khasso ykoun s7i7. Khdem `taxonomy planner` w 7tto f `listing.json`.
