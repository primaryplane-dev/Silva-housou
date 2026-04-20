# 包装出荷登録・包装出荷トレース Excel VBAシステム概要

## 1. システム概要

- **包装出荷登録**  
  出荷データの登録・管理を行う。  
  主に出荷予定・実績の入力、在庫引当、出荷先情報の管理などを担う。

- **包装出荷トレース**  
  出荷済み製品のトレース（追跡）や検索を行う。  
  出荷実績や在庫状況の参照、ロット・賞味期限等による検索が可能。

---

## 2. 主な使用テーブルと用途・カラム例

### ■ 包装出荷登録

| テーブル名         | 用途                     | 主なカラム例（物理名）                |
|--------------------|--------------------------|---------------------------------------|
| WNPP21B3           | 出荷先マスタ・受注予定   | JPTNO, JPSNO, JPKSU, JPSWK, JPHC4, JPDLT, JPSYS, JPPNE, JPPTU, JPPHI |
| SZSP01             | 出荷データ本体           | ZSSNO, ZSSGY, ZSSRY, ZSDLT, ZSYUCA, ZSLMT, ZSHNO, ZSLOT             |
| SSZP01             | 在庫引当ワーク           | SZSNO, SZDLT, SZLMT, SZSRY, SZSLD                                   |
| WTMP01             | 特約店マスタ（出荷先）   | TMTNO, TMKTM                                                          |
| WTEP01             | 特約店枝番管理マスタ     | TETNO, TEENO, TECD1, TEME1                                            |
| SRAP01             | 出荷期限ルールマスタ     | RATNO, RAENO, RACD1, RAPTN, RADLT                                    |
| WMSP01             | 運送会社マスタ           | MSMSC, MSMSK, MSRYM                                                   |
| BAEP01             | 外部マスタ参照           | AEANO, AEIKM など                                                     |

---

### ■ 包装出荷トレース

| テーブル名         | 用途                     | 主なカラム例（物理名）                |
|--------------------|--------------------------|---------------------------------------|
| SRHP01             | トレース本体             | RHHNO, RHDLT, LOT, SNO, HNM, LMT, SLD, SRY |
| SSZP01             | 在庫引当ワーク           | SZSNO, SZDLT, SZLMT, SZSRY, SZSLD     |
| SZSP01             | 出荷データ本体           | ZSSNO, ZSSGY, ZSSRY, ZSDLT, ZSYUCA, ZSLMT, ZSHNO, ZSLOT |
| WNPP21B3           | 出荷先マスタ・受注予定   | JPTNO, JPSNO, JPKSU, JPSWK, JPHC4, JPDLT, JPSYS, JPPNE, JPPTU, JPPHI |
| WTMP01             | 特約店マスタ（出荷先）   | TMTNO, TMKTM                         |
| WTEP01             | 特約店枝番管理マスタ     | TETNO, TEENO, TECD1, TEME1           |
| WSKP01             | 倉庫マスタ               | SKSKC, SKSKM                         |
| WMSP01             | 運送会社マスタ           | MSMSC, MSMSK, MSRYM                  |

---

