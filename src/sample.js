/**
 * Recon — sample data: an internal ledger vs. a custody statement for crypto
 * withdrawals and deposits, with planted breaks.
 */
var ReconSample = (function () {
  'use strict';
  function build(now) {
    now = now || new Date();
    var d = function (off) { return new Date(now.getFullYear(), now.getMonth(), now.getDate() + off); };
    var h = function (n) { return '0x' + (n * 2654435761 >>> 0).toString(16).padStart(8, '0') + 'a4f9c2' + String(n).padStart(4, '0'); };

    var ledger = [['Tx hash', 'Asset', 'Amount', 'Booked date', 'Wallet', 'Note']];
    var statement = [['Transaction ID', 'Coin', 'Quantity', 'Settled at', 'Fee']];
    var tx = [
      [1, 'BTC', 0.52, -14], [2, 'ETH', 12.4, -13], [3, 'USDC', 25000, -13], [4, 'BTC', -0.15, -12], [5, 'ETH', 3.1, -11],
      [6, 'USDC', -1200, -10], [7, 'BTC', 1.002, -9], [8, 'ETH', -0.75, -8], [9, 'USDC', 8800, -7], [10, 'BTC', 0.0375, -6],
      [11, 'ETH', 5.25, -5], [12, 'USDC', 15000, -4], [13, 'BTC', -0.3, -3], [14, 'ETH', 2.2, -2], [15, 'USDC', 4100, -1]
    ];
    tx.forEach(function (t) {
      var wallet = t[1] === 'BTC' ? 'Cold 1' : t[1] === 'ETH' ? 'Hot 2' : 'Hot 3';
      ledger.push([h(t[0]), t[1], t[2], d(t[3]), wallet, '']);
      statement.push([h(t[0]).toUpperCase().replace(/^0X/, ''), t[1], t[2], d(t[3] + (t[0] % 4 === 0 ? 1 : 0)), t[1] === 'USDC' ? 2 : 0.0001]);
    });

    // planted breaks
    statement[3][2] = 24990;                         // USDC amount differs by -10
    statement[5][3] = d(-11 + 3);                    // ETH settled 3 days later (tolerance 1)
    statement[8][1] = 'WETH';                        // asset differs
    statement.splice(10, 1);                         // tx 10 missing on the statement
    statement.push([h(99), 'ETH', 0.9, d(-2), 0.0001]);   // on the statement, never booked
    ledger.push([h(12), 'USDC', 15000, d(-4), 'Hot 3', 'booked twice']);   // duplicate booking
    statement.push(['', 'BTC', 0.25, d(-1), 0.0001]);     // custody row without ID ...
    ledger.push(['', 'BTC', 0.25, d(-1), 'Cold 1', 'manual entry']);       // ... matches a ledger row without ID

    return {
      a: { name: 'Internal ledger', rows: ledger },
      b: { name: 'Custody statement', rows: statement }
    };
  }
  return { build: build };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = ReconSample;
