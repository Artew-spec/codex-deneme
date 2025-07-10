document.addEventListener('DOMContentLoaded', function() {
  var paymentColumn = document.querySelector('.footer__column--info');
  if (!paymentColumn) return;

  var blocks = document.querySelectorAll('.footer__blocks-wrapper .footer-block--menu');
  blocks.forEach(function(block) {
    paymentColumn.appendChild(block);
  });
});
