from __future__ import annotations

import asyncio
from typing import Iterable, List, Dict, Any

from crawl4ai import async_webcrawler as c4
from crawl4ai.deep_crawling import BFSDeepCrawlStrategy


class PublicCrawler:
    def __init__(self, base_domain: str, out_dir: str) -> None:
        self.base_domain = base_domain.rstrip('/')
        self.out_dir = out_dir

    async def seed_urls(self, max_urls: int = 1000, hits_per_sec: int = 2) -> List[str]:
        crawler = c4.AsyncWebCrawler()
        seeding = c4.SeedingConfig(
            source="sitemap+cc",
            max_urls=max_urls,
            hits_per_sec=hits_per_sec,
            filter_nonsense_urls=True,
        )
        urls = await crawler.aseed_urls(self.base_domain, seeding)
        await crawler.close()
        # Normalize to list of strings
        merged: List[str] = []
        if isinstance(urls, dict):
            for _, values in urls.items():
                for v in values:
                    if isinstance(v, str):
                        merged.append(v)
                    elif isinstance(v, dict) and 'url' in v:
                        merged.append(v['url'])
        elif isinstance(urls, list):
            for v in urls:
                if isinstance(v, str):
                    merged.append(v)
                elif isinstance(v, dict) and 'url' in v:
                    merged.append(v['url'])
        return merged

    async def crawl_urls(self, urls: Iterable[str], concurrency: int = 8) -> List[Dict[str, Any]]:
        crawler = c4.AsyncWebCrawler(
            config=c4.BrowserConfig(headless=True),
        )

        run_cfg = c4.CrawlerRunConfig(
            check_robots_txt=True,
            deep_crawl_strategy=BFSDeepCrawlStrategy(max_depth=0),
            word_count_threshold=1,
            exclude_social_media_links=True,
            exclude_external_links=True,
            verbose=False,
        )

        results: List[Dict[str, Any]] = []
        async for item in crawler.arun_many(list(urls), config=run_cfg):
            # item is likely a CrawlResult or container entry
            try:
                md = getattr(item, 'markdown_v2', None) or getattr(item, 'markdown', '')
                url = getattr(item, 'url', '')
                title = getattr(item, 'title', '')
                if not md:
                    continue
                results.append({
                    'url': url,
                    'title': title,
                    'markdown': md,
                })
            except Exception:
                continue

        await crawler.close()
        return results


async def crawl_public(domain: str, limit: int, hits_per_sec: int = 2) -> List[Dict[str, Any]]:
    crawler = PublicCrawler(domain, out_dir="data/raw")
    urls = await crawler.seed_urls(max_urls=limit, hits_per_sec=hits_per_sec)
    # Keep only URLs within the registrable domain
    from urllib.parse import urlparse
    want_host = urlparse(domain).hostname or ""
    registrable = ".".join(want_host.split(".")[-2:]) if want_host else "test-achats.be"
    urls = [u for u in urls if urlparse(u).hostname and urlparse(u).hostname.endswith(registrable)]
    return await crawler.crawl_urls(urls)


def main() -> None:
    import argparse, json, os
    parser = argparse.ArgumentParser()
    parser.add_argument('--domain', required=True)
    parser.add_argument('--limit', type=int, default=200)
    parser.add_argument('--hits-per-sec', type=int, default=2)
    args = parser.parse_args()

    os.makedirs('data/raw', exist_ok=True)
    items = asyncio.run(crawl_public(args.domain, args.limit, args.hits_per_sec))
    out_path = 'data/raw/crawl.jsonl'
    with open(out_path, 'w', encoding='utf-8') as f:
        for it in items:
            f.write(json.dumps(it, ensure_ascii=False) + "\n")
    print(f"Wrote {len(items)} pages to {out_path}")


if __name__ == '__main__':
    main()

