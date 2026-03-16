function calcDaysAgo_(dateStr) {
  const KST = 'Asia/Seoul';
  const today = Utilities.formatDate(new Date(), KST, 'yyyy-MM-dd');
  if (dateStr === today) return '오늘';
  const diffMs = new Date(today) - new Date(dateStr);
  const days = Math.floor(diffMs / 86400000);
  if (days === 1) return '어제';
  return `${days}일 전`;
}
