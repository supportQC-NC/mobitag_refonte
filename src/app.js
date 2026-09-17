/* Mobitag — parcours d'envoi.
   Amélioration progressive uniquement : sans JavaScript, le formulaire
   reste un POST classique vers /envoyer, et la validation du numéro est
   assurée par l'attribut pattern. Chargé en defer depuis le head. */
(function () {
  'use strict';

  var msg      = document.getElementById('message');
  var jauge    = document.getElementById('jauge');
  var compteur = document.getElementById('compteur');
  var apercu   = document.getElementById('apercu-corps');
  var vide     = document.getElementById('apercu-vide');
  var sign     = document.getElementById('signature');
  var num      = document.getElementById('destinataire');
  var blocNum  = document.getElementById('bloc-numero');
  var errNum   = document.getElementById('erreur-destinataire');
  var form     = document.getElementById('formulaire');
  var confirm  = document.getElementById('confirmation');
  var titreCnf = document.getElementById('titre-confirmation');
  var redac    = document.getElementById('redaction');
  var etapes   = document.getElementById('etapes');
  var pro      = document.getElementById('orientation-pro');
  var motifPro = document.getElementById('motif-pro');
  var codeVue  = document.getElementById('code-origine');
  var LIMITE   = 150;

  /* ==========================================================
     1. COMPTAGE SMS RÉEL
     Un SMS ne compte pas en caractères JavaScript. L'alphabet
     GSM 03.38 permet 160 caractères par segment ; dès qu'un seul
     caractère en sort — un emoji, une espace insécable, certains
     guillemets — le message bascule en UCS-2 et tombe à 70.
     Quelques caractères (€ { } [ ] ~ ^ | \) occupent deux places
     même en GSM-7. C'est ce qui détermine le coût réel côté
     plateforme, donc la ligne « coût d'acheminement » du modèle.
     ========================================================== */
  var GSM_BASE = '@£$¥èéùìòÇ\nØø\rÅåΔ_ΦΓΛΩΠΨΣΘΞÆæßÉ !"#¤%&\'()*+,-./0123456789:;<=>?'
               + '¡ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÑÜ§¿abcdefghijklmnopqrstuvwxyzäöñüà';
  var GSM_EXT  = '\f^{}\\[~]|€';

  function mesurer(texte) {
    var unites = 0, gsm = true;
    for (var i = 0; i < texte.length; i++) {
      var c = texte[i];
      if (GSM_BASE.indexOf(c) !== -1)      { unites += 1; }
      else if (GSM_EXT.indexOf(c) !== -1)  { unites += 2; }
      else                                 { gsm = false; break; }
    }
    if (!gsm) {
      // UCS-2 : on compte en unités de code UTF-16, comme le fait le réseau.
      unites = texte.length;
    }
    var parSeg = gsm ? 160 : 70;
    var parSegMulti = gsm ? 153 : 67;
    var segments = unites === 0 ? 1
                 : (unites <= parSeg ? 1 : Math.ceil(unites / parSegMulti));
    return { gsm: gsm, unites: unites, segments: segments };
  }

  /* ==========================================================
     2. CODE D'ORIGINE
     Généré par message, jamais réutilisé. Alphabet sans I, O, 0
     ni 1 : le destinataire doit pouvoir le recopier sans se
     tromper, éventuellement en le lisant à voix haute.
     ========================================================== */
  var ALPHABET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';

  function nouveauCode() {
    var tirage = new Uint8Array(5);
    (window.crypto || window.msCrypto).getRandomValues(tirage);
    var code = '';
    for (var i = 0; i < 5; i++) { code += ALPHABET[tirage[i] % ALPHABET.length]; }
    return code;
  }

  var codeCourant = nouveauCode();
  codeVue.textContent = codeCourant;

  /* ==========================================================
     3. DÉTECTION D'UN USAGE PROFESSIONNEL — F10
     Faisceau d'indices, jamais bloquant : le message part quand
     même. On propose Helia PRO au moment où l'utilisateur se
     heurte à la limite du service grand public.
     ========================================================== */
  var INDICES = [
    [/\b(promo|promotion|soldes?|destockage|déstockage|remise exceptionnelle)\b/i,
     'Votre message ressemble à une offre commerciale.'],
    [/\b(devis|facture|bon de commande|acompte|règlement sous)\b/i,
     'Votre message évoque une relation commerciale.'],
    [/\b(nos horaires|notre boutique|notre magasin|votre commande|votre livraison|votre rendez-vous)\b/i,
     'Votre message s\'adresse à un client.'],
    [/\b(cher client|chers clients|chère cliente|bonjour à tous|à tous nos)\b/i,
     'Votre message semble destiné à plusieurs personnes.'],
    [/\b(SARL|SAS|SASU|EURL|SCI|EIRL)\b/,
     'Votre signature mentionne une société.'],
    [/https?:\/\/|www\./i,
     'Votre message contient un lien.'],
    [/\b(stop\s?sms|pour ne plus recevoir|désabonn)/i,
     'Votre message contient une mention de désabonnement, propre aux envois professionnels.']
  ];

  var proRefuse = false;

  function jaugerUsage() {
    if (proRefuse) { return; }
    var texte = msg.value + ' ' + sign.value;
    for (var i = 0; i < INDICES.length; i++) {
      if (INDICES[i][0].test(texte)) {
        motifPro.textContent = INDICES[i][1];
        pro.hidden = false;
        return;
      }
    }
    pro.hidden = true;
  }

  document.getElementById('refuser-pro').addEventListener('click', function () {
    proRefuse = true;
    pro.hidden = true;
    msg.focus();
  });

  /* ==========================================================
     4. RENDU
     ========================================================== */
  function rendre() {
    var texte = msg.value;
    var reste = LIMITE - texte.length;
    var m = mesurer(texte);

    jauge.style.setProperty('--reste', Math.max(0, reste / LIMITE * 100) + '%');
    jauge.dataset.niveau = reste <= 0 ? 'plein' : (reste <= 25 ? 'tendu' : '');

    var libelle;
    if (reste > 0) {
      libelle = reste + ' caractère' + (reste > 1 ? 's' : '')
              + ' disponible' + (reste > 1 ? 's' : '');
    } else {
      libelle = 'Limite atteinte';
    }
    if (texte.length > 0) {
      libelle += ' · ' + m.segments + ' SMS';
      if (!m.gsm) { libelle += ' (alphabet étendu : un emoji ou un caractère spécial réduit la capacité)'; }
    }
    compteur.textContent = libelle;

    var suffixe = sign.value.trim() ? '\n— ' + sign.value.trim() : '';
    if (texte) {
      apercu.textContent = texte + suffixe;   // textContent : aucune injection possible
    } else {
      apercu.textContent = '';
      apercu.appendChild(vide);
    }
    jaugerUsage();
  }

  msg.addEventListener('input', rendre);
  sign.addEventListener('input', rendre);

  /* Mise en forme du numéro pendant la saisie : 75 12 34 */
  num.addEventListener('input', function () {
    var chiffres = num.value.replace(/\D/g, '').slice(0, 6);
    num.value = chiffres.replace(/(\d{2})(?=\d)/g, '$1 ').trim();
    verifierNumero(false);
  });

  function verifierNumero(strict) {
    var chiffres = num.value.replace(/\D/g, '');
    var message = '';
    if (chiffres.length === 0)            { message = strict ? 'Indiquez le numéro du destinataire.' : ''; }
    else if (!/^[789]/.test(chiffres))    { message = 'Ce numéro n\'est pas un mobile calédonien. Les mobiles Mobilis et Liberté commencent par 7, 8 ou 9.'; }
    else if (chiffres.length < 6)         { message = strict ? 'Un numéro mobile compte six chiffres.' : ''; }

    errNum.textContent = message;
    errNum.hidden = !message;
    blocNum.classList.toggle('invalide', !!message);
    return !message && chiffres.length === 6;
  }
  num.addEventListener('blur', function () { verifierNumero(true); });

  /* ==========================================================
     5. CYCLE DE VIE — F09
     Le simulateur est déterministe, pour qu'une démonstration
     soit reproductible devant un jury :
       99 xx xx  → échec de remise (mobile injoignable)
       tout autre → remise normale
     ========================================================== */
  var CYCLE_OK = [
    ['Accepté',            'Le message a passé les contrôles.'],
    ['En file d\'attente', 'Transmis à la plateforme SMS.'],
    ['Remis',              'Reçu sur le mobile du destinataire.']
  ];
  var CYCLE_KO = [
    ['Accepté',            'Le message a passé les contrôles.'],
    ['En file d\'attente', 'Transmis à la plateforme SMS.'],
    ['Non remis',          'Le mobile du destinataire est injoignable. Votre message n\'a pas été perdu.']
  ];

  function heure() {
    var t = new Date();
    return String(t.getHours()).padStart(2, '0') + ':' +
           String(t.getMinutes()).padStart(2, '0') + ':' +
           String(t.getSeconds()).padStart(2, '0');
  }

  function ligneEtape(etape, fait, echec) {
    var li = document.createElement('li');
    var h  = document.createElement('span');
    h.className = 'heure';
    h.textContent = fait ? heure() : '—';
    var corps = document.createElement('span');
    if (fait) {
      var b = document.createElement('strong');
      b.textContent = etape[0];
      if (echec) { b.style.color = 'var(--alerte)'; }
      corps.appendChild(b);
      corps.appendChild(document.createTextNode(' — ' + etape[1]));
    } else {
      corps.className = 'attente';
      corps.textContent = etape[0] + '…';
    }
    li.appendChild(h);
    li.appendChild(corps);
    return li;
  }

  function envoyer() {
    var chiffres = num.value.replace(/\D/g, '');
    var echoue   = /^99/.test(chiffres);
    var cycle    = echoue ? CYCLE_KO : CYCLE_OK;

    document.getElementById('recap-numero').textContent = '+687 ' + num.value;
    document.getElementById('recap-phrase').hidden = false;
    document.getElementById('recap-code').hidden   = true;
    document.getElementById('bloc-brouillon').hidden = true;
    document.getElementById('reessayer').hidden    = true;
    confirm.classList.remove('echec');
    titreCnf.textContent = 'Envoi en cours';
    etapes.textContent = '';

    redac.hidden   = true;
    confirm.hidden = false;
    titreCnf.focus();               // le focus suit l'utilisateur

    cycle.forEach(function (etape, i) {
      var li = ligneEtape(etape, false, false);
      etapes.appendChild(li);
      var dernier = (i === cycle.length - 1);

      setTimeout(function () {
        etapes.replaceChild(ligneEtape(etape, true, echoue && dernier), li);
        if (!dernier) { return; }

        if (echoue) {
          titreCnf.textContent = 'Message non remis';
          confirm.classList.add('echec');
          document.getElementById('recap-phrase').hidden = true;
          document.getElementById('bloc-brouillon').hidden = false;
          document.getElementById('reessayer').hidden = false;
        } else {
          titreCnf.textContent = 'Message remis';
          document.getElementById('code-remis').textContent = codeCourant;
          document.getElementById('recap-code').hidden = false;
        }
      }, 700 + i * 1100);
    });
  }

  form.addEventListener('submit', function (e) {
    e.preventDefault();
    if (!verifierNumero(true)) { num.focus(); return; }
    if (!form.reportValidity()) { return; }
    envoyer();
  });

  /* Réessai : le message est toujours dans le formulaire, rien n'est ressaisi. */
  document.getElementById('reessayer').addEventListener('click', envoyer);

  document.getElementById('retour').addEventListener('click', function () {
    form.reset();
    proRefuse = false;
    codeCourant = nouveauCode();     // un message, un code
    codeVue.textContent = codeCourant;
    verifierNumero(false);
    rendre();
    confirm.hidden = true;
    redac.hidden   = false;
    msg.focus();
  });

  rendre();
})();
